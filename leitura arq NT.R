# leitura arquivo NET interface IOS

library(readxl)
library(readr)
library(stringr)
library(dplyr)
library(tidyr)


lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "BR2.*.NT.*.csv")
lista_arquivos

rm(nt)
for (arq in lista_arquivos) {
  temp <- read.csv2(arq, sep = ";", col.names = c("Mes", "DtFim", "Remito", "ArtigoSAP","Qtde","CustoTot","Empresa","Agrup"))
  temp <- temp[, -c(2,3,8)]
  if (exists("nt")) {
    nt <- rbind(nt, temp)
  }
  else {
    nt <- temp
  }
}

# case_when, alternativa para iselse
nt$Mes <- substr(nt$Mes, 1, 6)
nt$Empresa <- case_when(nt$Empresa == 233 ~ "DDB", nt$Empresa == 277 ~ "DLF")
colSums(is.na(nt))
nt$CustoTot <- as.numeric(nt$CustoTot)

nt <- aggregate(cbind(nt$Qtde, nt$CustoTot), list(nt$Mes, nt$Empresa, nt$ArtigoSAP), FUN = sum)
names(nt) <- c("Mes", "Empresa", "ArtigoSAP", "QtdeNT", "CustoNT")
View(nt)


mov_gamma_mes <- read.csv('mov_gamma_mes.csv', sep = ";")
View(mov_gamma_mes)
colSums(is.na(mov_gamma_mes))
mov_gamma_mes <- mov_gamma_mes[mov_gamma_mes$Tester.GWP == "N", ]

# Agregando por empresa. Removendo UN.
mov_gamma_mes <- aggregate(mov_gamma_mes[ , c(7:17)],mov_gamma_mes[ , c(1,3,4,5,6)], FUN = sum)

mov_mes <- merge(mov_gamma_mes, nt, by = c("Empresa", "ArtigoSAP"), all = TRUE)
View(mov_mes)

mov_mes$QtdeNTCalculado <- mov_mes$Saida.POS.SALES_REMITO + mov_mes$Saida.INVEN.SHORTAGES_REMITO
mov_mes$Mes[is.na(mov_mes$Mes)] <- nt$Mes[1]
mov_mes$QtdeNT[is.na(mov_mes$QtdeNT)] <- 0
mov_mes$CustoNT[is.na(mov_mes$CustoNT)] <- 0

write_csv2(mov_mes, str_replace('mov_mes_MESREF.csv', 'MESREF', nt$Mes[1]))
