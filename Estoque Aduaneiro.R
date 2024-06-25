library(readxl)
library(readr)
library(stringr)
library(dplyr)
library(RMySQL)

### ADUANEIRO ###
# INPEREBF
mes_ref <- paste(substr(aduaneiro$`Fecha proceso`, 7, 4), substr(aduaneiro$`Fecha proceso`, 4, 2), sep = "")

lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INPEREBF.*.xlsx")
lista_arquivos

rm(aduaneiro)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq,col_types = "text")
  #temp <- temp[temp$A) == "5285629" , ]
  temp <- temp[!is.na(temp$`Stock disponible`) , ]
  if (exists("aduaneiro")) {
    aduaneiro <- rbind(aduaneiro, temp)
  }
  else {
    aduaneiro <- temp
  }
}

# somente fakes
#aduaneiro <- aduaneiro[aduaneiro$`REMITO IOS` %in% c(666666, 777777, 999999) & !is.na(aduaneiro$`ARTIGO SAP`), ]
#aduaneiro <- aduaneiro[!is.na(aduaneiro$`PAGO IOS`) & !is.na(aduaneiro$`ARTIGO SAP`) & aduaneiro$`TTG` == 'NO', ]

colSums(is.na(aduaneiro))
str(aduaneiro)
View(aduaneiro)
sum(aduaneiro$`Stock disponible`)
sum(aduaneiro$`Saldo disponible calculado`)

#salva resultado
write_csv2(aduaneiro, str_replace('aduaneiro_MESREF.csv', 'MESREF', mes_ref))

# consolidação artigo / empresa / TTG
aduaneiro_artigo <- aggregate(aduaneiro$QTDE. , list(aduaneiro$UN, aduaneiro$TTG, aduaneiro$`ARTIGO GAMMA`, aduaneiro$`ARTIGO SAP`), FUN = sum)
names(aduaneiro_artigo) <- c("Empresa" ,"TTG", "ArtigoGamma", "ArtigoSAP", "Qtde")
View(aduaneiro_artigo)

# consolidação artigo / empresa 
aduaneiro_artigo_sobras <- aggregate(aduaneiro$QTDE. , list(aduaneiro$UN, aduaneiro$`ARTIGO GAMMA`, aduaneiro$`ARTIGO SAP`), FUN = sum)
names(aduaneiro_artigo_sobras) <- c("Empresa" ,"ArtigoGamma", "ArtigoSAP", "Qtde")
View(aduaneiro_artigo_sobras)

# Salvando resultado
write_csv2(aduaneiro, str_replace('aduaneiro_MESREF.csv', 'MESREF', mes_ref))
write_csv2(aduaneiro_artigo, str_replace('aduaneiro_artigo_MESREF.csv', 'MESREF', mes_ref))
write_csv2(aduaneiro_artigo_sobras, str_replace('aduaneiro_sobras_artigo_MESREF.csv', 'MESREF', mes_ref))
