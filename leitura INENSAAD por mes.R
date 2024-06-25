library(readxl)
library(writexl)
library(readr)
library(stringr)
library(dplyr)
library(tidyr)

# ENTRADAS

lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INENSAAD_ENTRADA.*.xlsx")
lista_arquivos

rm(entradas)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq,col_types = "text")
  temp <- temp[!is.na(temp$`Cantidad recibida`), c(1, 2, 3, 4, 6, 9, 10, 13, 16, 18, 19, 21)]
  if (exists("entradas")) {
    entradas <- rbind(entradas, temp)
  }
  else {
    entradas <- temp
  }
}

entradas$`Cantidad recibida` <- as.numeric(entradas$`Cantidad recibida`) 
entradas$`Coste unitario entrada` <- as.numeric(entradas$`Coste unitario entrada`) 
entradas$`Tester/GWP` <- ifelse(is.na(entradas$`Tester/GWP`) , "N", "Y") 
entradas$`Codigo centro receptor` <- substr(entradas$`Codigo centro receptor`, 1, 2)
entradas$`Codigo Aeropuerto` <- ifelse(entradas$`Codigo Aeropuerto` == "URG", "DLF", entradas$`Codigo Aeropuerto`)
entradas$tipo_entrada <- ifelse(entradas$`Numero factura` == "666666", "SOBRAS", ifelse( grepl("B", entradas$`Numero factura`),  'INTERCOMPANY', 'REMITO'))
View(entradas)

entradas$mes_ref <- as.Date(format(as.Date(as.numeric(entradas$`Fecha entrada`), origin = '1899-12-30'), "%Y-%m-01"))
entradas$`Fecha entrada` <- NULL

unique(entradas$tipo_entrada)

write_csv2(entradas, 'entradas_mes.csv')

sum(entradas$`Cantidad recibida`)
# removendo entradas de sobras que não são controladas pela IOS
#entradas <- entradas[entradas$`Numero factura` != "666666", ]

# totalizando por empresa / artigo
entradas_art <- aggregate(entradas$`Cantidad recibida`, list(entradas$`Codigo Aeropuerto`, entradas$`Codigo centro receptor`, entradas$mes_ref, entradas$`Articulo codigo Gamma`, entradas$`Articulo codigo SAP`, entradas$`Tester/GWP`, entradas$`Codigo Categoria`, entradas$tipo_entrada), FUN = sum)
names(entradas_art) <- c("Empresa" , "UN", "Mes", "ArtigoGamma", "ArtigoSAP", "Tester/GWP", "Categoria", "TipoEntrada" , "QtdeTot")
View(entradas_art)
sum(entradas_art$QtdeTot)
write_csv2(entradas_art, 'entradas_artigo.csv')

# SAIDA
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INENSAAD_SAIDA.*.xlsx")
lista_arquivos

rm(saidas)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq,col_types = "text")
  temp <- temp[!is.na(temp$`Cantidad Salida`), c(1, 2, 3, 4, 6, 9, 10, 11, 12,14, 15, 18, 19)]
  if (exists("saidas")) {
    saidas <- rbind(saidas, temp)
  }
  else {
    saidas <- temp
  }
}

saidas$`Cantidad Salida` <- as.numeric(saidas$`Cantidad Salida`)
saidas$`Coste unitario entrada` <- as.numeric(saidas$`Coste unitario entrada`)
saidas$`Tester/GWP` <- ifelse(is.na(saidas$`Tester/GWP`), "N", "Y")
saidas$`Codigo Centro Salida` <- substr(saidas$`Codigo Centro Salida`,1,2)
saidas$`Codigo Aeropuerto` <- ifelse(saidas$`Codigo Aeropuerto` == "URG", "DLF", saidas$`Codigo Aeropuerto`)
saidas$IndSobras <- ifelse(is.na(saidas$`Numero factura original`) | saidas$`Numero factura original` %in% c("666666", "777777", "999999"), "SOBRAS", "REMITO") 
View(saidas)
sum(saidas$`Cantidad Salida`)

saidas$mes_ref <- as.Date(format(as.Date(as.numeric(saidas$`Fecha Salida`), origin = '1899-12-30'), "%Y-%m-01"))
saidas$`Fecha Salida` <- NULL

write_csv2(saidas, 'saidas_mes.csv')

# totalizando por empresa / artigo
saidas_art <- aggregate(saidas$`Cantidad Salida`, list(saidas$`Codigo Aeropuerto`, saidas$`Codigo Centro Salida`, saidas$mes_ref, saidas$`Articulo codigo Gamma`, saidas$`Articulo codigo SAP`, saidas$`Tester/GWP`, saidas$`Codigo Categoria`, saidas$`Tipo Salida`, saidas$IndSobras), FUN = sum)
names(saidas_art) <- c("Empresa" , "UN", "Mes", "ArtigoGamma", "ArtigoSAP", "Tester/GWP", "Categoria", "TipoSaida" , "IndSobra", "QtdeTot")
View(saidas_art)
sum(saidas_art$QtdeTot)

write_csv2(saidas_art, 'saidas_artigo.csv')



# visão Gamma movimentação mensal
entradas_art <- read.csv('entradas_artigo.csv', sep = ";")
saidas_art <- read.csv('saidas_artigo.csv', sep = ";")

View(entradas_art)
View(saidas_art)

entradas_art$TipoEntrada <- paste("Entrada.", entradas_art$TipoEntrada, sep = "")
mov <- pivot_wider(entradas_art, names_from = TipoEntrada, values_from = QtdeTot, values_fill = 0)
View(mov)

saidas_art$TipoSaida <- paste("Saida.", saidas_art$TipoSaida, sep = "")
temp <-pivot_wider(saidas_art, names_from = c(TipoSaida, IndSobra), values_from = QtdeTot, values_fill = 0)
names(temp)
View(temp)


mov_gamma <- merge(mov, temp, by = c("Empresa", "UN", "Mes", "ArtigoGamma", "ArtigoSAP", "Tester/GWP", "Categoria"), all = TRUE)
mov_gamma[is.na(mov_gamma)] <- 0
View(mov_gamma)

write_xlsx(mov_gamma, 'mov_gamma_mes.xlsx')


