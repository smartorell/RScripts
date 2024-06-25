library(readxl)
library(readr)
library(stringr)
library(dplyr)
library(RMySQL)
library(writexl)


# Estoque comercial - Vegga
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "Estoque Vegga DP.*.xlsx")
lista_arquivos

rm(comercial)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq)
  temp <- temp[temp$Qtde_Total != 0 , c(1,2,4,9,14,15,18, 19)]
  if (exists("comercial")) {
    comercial <- rbind(comercial, temp)
  }
  else {
    comercial <- temp
  }
}

colSums(is.na(comercial))
View(comercial)
str(comercial)
comercial$Consignado <- ifelse(comercial$`On Consignment Units` > 0, 'S', 'N')

mes_ref <- paste(substr(comercial$DATE[1], 4, 7), substr(comercial$DATE[1], 1, 2), sep = "")

# consolidação artigo / empresa
comercial_artigo <- aggregate(comercial$Qtde_Total, list(comercial$`Unidade Negocio`, comercial$`Shop Code`, comercial$`Gamma Product Code`, comercial$Consignado), FUN = sum)
names(comercial_artigo) <- c("Empresa", "Centro" , "ArtigoGamma", "Consignado", "Qtde")
View(comercial_artigo)
sum(comercial_artigo$Qtde)

# ADUANEIRO
   # INPERENT
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INPERENT.*.xlsx")
lista_arquivos

rm(aduaneiro)
for (arq in lista_arquivos) {
     temp <- read_excel(path = arq,col_types = "text")
     temp <- temp[!is.na(temp$`Dispnibl`) , c(2,5, 8,12,14, 20)]
     if (exists("aduaneiro")) {
       aduaneiro <- rbind(aduaneiro, temp)
     }
     else {
       aduaneiro <- temp
     }
  }


colSums(is.na(aduaneiro))
aduaneiro$`Dispnibl` <- as.numeric(aduaneiro$`Dispnibl`)
aduaneiro$`Fec. Entrada` <- as.numeric(aduaneiro$`Fec. Entrada`)
aduaneiro <- aduaneiro[aduaneiro$`Fec. Entrada` < 20231200, ]
str(aduaneiro)
View(aduaneiro)
sum(aduaneiro$Dispnibl)

# consolidação artigo / empresa
aduaneiro_artigo <- aggregate(aduaneiro$`Dispnibl`, list(aduaneiro$Centro, aduaneiro$Artículo), FUN = sum)
names(aduaneiro_artigo) <- c("Centro" , "ArtigoGamma", "Qtde")
View(aduaneiro_artigo)
sum(aduaneiro_artigo$Qtde)

# Salvando resultado
write_csv2(comercial, str_replace('comercial_MESREF.csv', 'MESREF', mes_ref))
write_csv2(comercial_artigo, str_replace('comercial_artigo_MESREF.csv', 'MESREF', mes_ref))
write_csv2(aduaneiro, str_replace('aduaneiro_MESREF.csv', 'MESREF', mes_ref))
write_csv2(aduaneiro_artigo, str_replace('aduaneiro_artigo_MESREF.csv', 'MESREF', mes_ref))

# Comparacao
comparacao <- merge(comercial_artigo, aduaneiro_artigo, by = c("Centro", "ArtigoGamma"), all = TRUE)
View(comparacao)
comparacao <- rename(comparacao, Qtde_comercial = Qtde.x, Qtde_aduaneiro = Qtde.y)
colSums(is.na(comparacao))
comparacao[is.na(comparacao)] <- 0

comparacao$Qtde_comercial <- ifelse(comparacao$Qtde_comercial < 0, 0, comparacao$Qtde_comercial)
diferenca <- comparacao[comparacao$Qtde_comercial - comparacao$Qtde_aduaneiro != 0 & comparacao$Qtde_comercial >= 0, ]
diferenca$diff <- diferenca$Qtde_aduaneiro - diferenca$Qtde_comercial
View(diferenca)

diferenca <- arrange(diferenca, ArtigoGamma)

write_xlsx(diferenca, str_replace('diferenca_DP_MESREF.csv', 'MESREF', mes_ref))

sum(comparacao$Qtde_comercial)
sum(comparacao$Qtde_aduaneiro)
