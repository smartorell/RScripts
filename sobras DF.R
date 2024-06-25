library(readxl)
library(readr)
library(stringr)
library(dplyr)
library(RMySQL)
library(lubridate)

lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INEAFRET.*.xlsx")
lista_arquivos

rm(aduaneiro)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq,col_types = "text")
  temp <- temp[(!is.na(temp$QTDE.) & temp$TTG == 'NO') , c(2,3, 11, 12, 14, 17, 19, 23)]
  if (exists("aduaneiro")) {
    aduaneiro <- rbind(aduaneiro, temp)
  }
  else {
    aduaneiro <- temp
  }
}

aduaneiro$`PAGO IOS`[is.na(aduaneiro$`PAGO IOS`)] <- 'N'
aduaneiro <- aduaneiro[aduaneiro$`PAGO IOS` == 'S', -c(5,8)]

colSums(is.na(aduaneiro))
sobras <- aduaneiro
sobras$Empresa <- ifelse(sobras$UN == 'BD', 'DDB', 'DLF')
sobras$QTDE. <- as.numeric(sobras$QTDE.)
sobras$`FOB TOTAL` <- as.numeric(sobras$`FOB TOTAL`)
sum(sobras$QTDE.)
colSums(is.na(sobras))
str(sobras)
View(sobras)

# removendo linhas sem artigo SAP
sobras <- sobras[!is.na(sobras$`ARTIGO SAP`), ]
sobras <- sobras[sobras$QTDE. > 0, ]

sobras$UN <- NULL

sobras <- aggregate(cbind(sobras$QTDE., sobras$`FOB TOTAL`) , list(sobras$Empresa, sobras$CAT., sobras$`ARTIGO GAMMA`, sobras$`ARTIGO SAP`), FUN = sum)
names(sobras) <- c("Empresa", "Categoria", "Artigo Gamma", "Artigo SAP", "Qty", "Vl Total")
sum(sobras$Qty)

write_csv2(sobras, 'sobras.csv')
