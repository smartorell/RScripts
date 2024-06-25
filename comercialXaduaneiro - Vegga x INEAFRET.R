library(readxl)
library(readr)
library(stringr)
library(dplyr)
library(RMySQL)
library(lubridate)
library(writexl)


# CORREÇÕES A FAZER: Estoque vegga em xlsx está com artigo com letras de notação cientifica e primeira coluna está com formato de data diferente de MON/YY


### Estoque comercial - Vegga ###
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "Estoque Vegga DF.*.xlsx")
lista_arquivos

rm(comercial)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq)
  #temp <- temp[temp$Qtde_Total != 0 , c(1,4,14,15,18, 19)]
  temp <- temp[temp$`On Consignment Units` != 0 , c(1, 4,9,14, 16,18)]
  if (exists("comercial")) {
    comercial <- rbind(comercial, temp)
  }
  else {
    comercial <- temp
  }
}

View(comercial)
colSums(is.na(comercial))

dim(comercial)
str(comercial)

#mes_ref <- paste(toupper(substr(format(comercial$DATE[1], format = "%B"), 1, 3)), format(comercial$DATE[1], format = "%y"), sep = "")
mes_ref <- paste(substr(comercial$DATE[1], 4, 7), substr(comercial$DATE[1], 1, 2), sep = "")


# consolidação artigo / empresa
comercial_artigo <- aggregate(comercial$`On Consignment Units`, list(comercial$`Unidade Negocio`, comercial$`Global Category Code`, comercial$`Gamma Product Code`), FUN = sum)
names(comercial_artigo) <- c("Empresa" ,"Categoria", "ArtigoGamma", "Qtde")
View(comercial_artigo)

sum(comercial$`On Consignment Units`)
sum(comercial_artigo$Qtde)

### ADUANEIRO ###

lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INEAFRET.*.xlsx")
lista_arquivos

rm(aduaneiro)
for (arq in lista_arquivos) {
     temp <- read_excel(path = arq,col_types = c("skip", "text", "text", "skip", "skip", "skip", "skip", "skip", "skip", "skip", "text", "text", "skip", "text", "date", "text", "numeric", "skip", "skip", "skip", "skip", "skip", "skip", "skip", "skip", "skip", "skip"))
     temp <- temp[!is.na(temp$QTDE.) , ] #c(2,3, 11, 12, 14,15, 16, 17)]
     if (exists("aduaneiro")) {
       aduaneiro <- rbind(aduaneiro, temp)
     }
     else {
       aduaneiro <- temp
     }
  }


colSums(is.na(aduaneiro))
#aduaneiro$QTDE. <- as.numeric(aduaneiro$QTDE.)
str(aduaneiro)
View(aduaneiro)
# aduaneiro <- aduaneiro[aduaneiro$DT.ENTRADA < (comercial$DATE[1] %m+% months(1)), -c(6)]
aduaneiro <- aduaneiro[aduaneiro$DT.ENTRADA < as.Date(paste(mes_ref, "01", sep = ""), format = "%Y%m%d") %m+% months(1), -c(6)]


# consolidação artigo / empresa
aduaneiro_artigo <- aggregate(aduaneiro$QTDE. , list(aduaneiro$UN, aduaneiro$CAT., aduaneiro$`ARTIGO GAMMA`), FUN = sum)
names(aduaneiro_artigo) <- c("Empresa" ,"Categoria", "ArtigoGamma", "Qtde")
View(aduaneiro_artigo)

sum(aduaneiro_artigo$Qtde)
sum(aduaneiro$QTDE.)


# Salvando resultado
write_csv2(comercial, str_replace('comercial_MESREF.csv', 'MESREF', mes_ref))
write_csv2(comercial_artigo, str_replace('comercial_artigo_MESREF.csv', 'MESREF', mes_ref))
write_csv2(aduaneiro, str_replace('aduaneiro_MESREF.csv', 'MESREF', mes_ref))
write_csv2(aduaneiro_artigo, str_replace('aduaneiro_artigo_MESREF.csv', 'MESREF', mes_ref))

# Comparacao
#comercial_artigo <- read.csv('comercial_artigo_ABR23.csv', sep = ";")
#View(comercial_artigo)

comparacao <- merge(comercial_artigo, aduaneiro_artigo, by = c("Empresa", "ArtigoGamma"), all = TRUE)
View(comparacao)
comparacao$Categoria.x <- ifelse(is.na(comparacao$Categoria.x), comparacao$Categoria.y, comparacao$Categoria.x ) 
comparacao$Categoria.y <- NULL
comparacao <- rename(comparacao, Categoria = Categoria.x, Qtde_comercial = Qtde.x, Qtde_aduaneiro = Qtde.y)
colSums(is.na(comparacao))
comparacao[is.na(comparacao)] <- 0

sum(comparacao$Qtde_comercial)
comparacao$Qtde_comercial[comparacao$Qtde_comercial < 0] <- 0

sum(comparacao$Qtde_aduaneiro)

#comparacao <- read.csv('comparacao_ABR23.csv', sep = ";")
View(comparacao)

#Pegando transito do relatorio vegga
transito <- aggregate(comercial$`Transit Units`, list(comercial$`Unidade Negocio`, comercial$`Gamma Product Code`), FUN = sum)

# Pegando transito do INFALTRA (opção usada atualmente)
transito <- read_excel(list.files(getwd(), full.names = TRUE, pattern  = "INFALTRA.*.xlsx")[1])
View(transito)
transito <- transito[transito$`Tránsito Comercial` == "S" & transito$`Tránsito Aduanero` == "N", ]
transito <- aggregate(transito$Cantidad, list(substr(transito$`Centro Origen`,1,2), transito$`Código Artículo Gamma`), FUN = sum)
names(transito) <- c("Empresa","ArtigoGamma","QtdeInTransit")
#

comparacao <- merge(comparacao, transito, by = c("Empresa", "ArtigoGamma"), all = TRUE)
comparacao$QtdeInTransit[is.na(comparacao$QtdeInTransit)] <- 0
comparacao$QtdeComercialTot <- comparacao$Qtde_comercial + comparacao$QtdeInTransit

comparacao$diff_stran <- comparacao$Qtde_aduaneiro - comparacao$Qtde_comercial
comparacao$diff_ctran <- comparacao$Qtde_aduaneiro - comparacao$QtdeComercialTot

#calculando diferenca
diferenca <- comparacao[comparacao$diff_stran != 0 & comparacao$diff_ctran != 0 , ]
#diferenca$diff = diferenca$Qtde_aduaneiro - diferenca$Qtde_comercial
diferenca <- arrange(diferenca, ArtigoGamma)
View(diferenca)

diferenca <- arrange(diferenca, desc(abs(diff_ctran + diff_stran)))

sum(diferenca$diff_ctran)

# Salvando resultado em CSV
write_csv2(comparacao, str_replace('comparacao_MESREF.csv', 'MESREF', mes_ref))
#write_csv2(diferenca, str_replace('diferenca_MESREF.csv', 'MESREF', mes_ref))
write_xlsx(diferenca, str_replace('diferenca_MESREF.xlsx', 'MESREF', mes_ref))



