library(readxl)
library(readr)
library(stringr)
library(dplyr)
library(RMySQL)


# Estoque comercial - Vegga
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "Estoque Vegga DF.*.xlsx")
lista_arquivos

rm(comercial)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq)
  temp <- temp[temp$`On Consignment Units` != 0 , c(1,9,10,12,13,14,15)]
  if (exists("comercial")) {
    comercial <- rbind(comercial, temp)
  }
  else {
    comercial <- temp
  }
}

colSums(is.na(comercial))
View(comercial)
unique(comercial$Tester)
str(comercial)

mes_ref <- "FEV23" #melhorar

# consolidação artigo / empresa
comercial_artigo <- aggregate(comercial$`On Consignment Units`, list(comercial$`Unidade Negocio`, comercial$`Product Code`), FUN = sum)
names(comercial_artigo) <- c("Empresa" , "ArtigoGamma", "Qtde")
View(comercial_artigo)

# ADUANEIRO
   # INPEREBF
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INPEREBF.*.xlsx")
lista_arquivos

rm(aduaneiro)
for (arq in lista_arquivos) {
     temp <- read_excel(path = arq,col_types = "text")
     temp <- temp[!is.na(temp$`Stock disponible`) & is.na(temp$Observaciones), c(2, 10, 12, 13, 20, 26)]
     if (exists("aduaneiro")) {
       aduaneiro <- rbind(aduaneiro, temp)
     }
     else {
       aduaneiro <- temp
     }
  }


colSums(is.na(aduaneiro))
aduaneiro$`Stock disponible` <- as.numeric(aduaneiro$`Stock disponible`)
str(aduaneiro)

aduaneiro$Deposito <- str_replace_all(aduaneiro$Deposito, c("DA27197888" = "BD", "DA17625216" = "BR", "FR17625216" = "BF" ))
View(aduaneiro)

# consolidação artigo / empresa
aduaneiro_artigo <- aggregate(aduaneiro$`Stock disponible`, list(aduaneiro$Deposito, aduaneiro$Articulo), FUN = sum)
names(aduaneiro_artigo) <- c("Empresa" , "ArtigoGamma", "Qtde")
View(aduaneiro_artigo)


# Salvando resultado
write_csv2(comercial, 'comercial_FEV23.csv')
write_csv2(comercial_artigo, 'comercial_artigo_FEV23.csv')
write_csv2(aduaneiro, 'aduaneiro_FEV23.csv')
write_csv2(aduaneiro_artigo, 'aduaneiro_artigo_FEV23.csv')

# Comparacao
comparacao <- merge(comercial_artigo, aduaneiro_artigo, by = c("Empresa", "ArtigoGamma"), all = TRUE)
View(comparacao)
comparacao <- rename(comparacao, Qtde_comercial = Qtde.x, Qtde_aduaneiro = Qtde.y)
colSums(is.na(comparacao))
comparacao[is.na(comparacao)] <- 0

diferenca <- comparacao[comparacao$Qtde_comercial - comparacao$Qtde_aduaneiro != 0 & comparacao$Qtde_comercial >= 0, ]
View(diferenca)
