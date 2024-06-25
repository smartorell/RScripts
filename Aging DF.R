library(readxl)
library(writexl)
library(readr)
library(stringr)
library(dplyr)
library(RMySQL)
library(lubridate)
library(tidyr)
# library(tidyverse)


# Leitura INPEREBF
# Notar que BF não deve ser considerado
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "INPEREBF.*.xlsx")
lista_arquivos

rm(base)
for (arq in lista_arquivos) {
  temp <- read_excel(path = arq, col_types = "text")
  temp <- temp[!is.na(temp$`Saldo disponible calculado`) & temp$`Saldo disponible calculado` > 0, -c(grep("\\d", colnames(temp)))]
  saldo_temp <- temp[ , c(13, 29:length(colnames(temp)))]
  temp <- temp[ , c(1,2,4, 7, 9, 10, 11, 12, 13, 16, 25, 26)] # incluida marca posição 12
  if (exists("base")) {
    base <- rbind(base, temp)
    saldo_aerop <- merge(saldo_aerop, saldo_temp, by = c("Articulo"), all = TRUE)
  }
  else {
    base <- temp
    saldo_aerop <- saldo_temp
  }
}

aeroportos <- colnames(saldo_aerop)
aeroportos <- aeroportos[-1]
saldo_aerop[is.na(saldo_aerop)] <- 0
saldo_aerop <- unique(saldo_aerop)
saldo_aerop <- rename(saldo_aerop, ITEM_CODE = Articulo)
# Transformar para numeric todas as colunas exceto a primeira
saldo_aerop[ , -1] <- sapply(saldo_aerop[ , -1], as.numeric)

saldo_aerop <- pivot_longer(saldo_aerop, cols = !ITEM_CODE, names_to = 'AIRPORT', values_to =  'Saldo')
saldo_aerop <- saldo_aerop %>% group_by(ITEM_CODE,AIRPORT) %>%
    summarise(Saldo = max(Saldo))

View(saldo_aerop)

dt_base <- as.Date(format(as.Date(as.numeric(base$`Fecha proceso`[1]), origin = '1899-12-30'), "%Y-%m-01"))
mes_ref <- format(dt_base - 1, "%Y%m")

# Essa parte deverá ser substituida por consulta direto no banco de dados
# recupera artigos novos (primeira entrada anterior a 1 ano)
novo <- read_excel('ITEM_NOVO.xlsx', col_types = "text")
View(novo)
novo$mes_entrada <-as.Date(novo$PRIMEIRA_ENTRADA, format = "%d/%m/%Y")
novo$meses <- interval(novo$mes_entrada, dt_base) %/% months(1)
novo$PRIMEIRA_ENTRADA <- NULL
novo$meses[novo$meses == 0] <- 1
# Recuperando venda média por aeroporto
vm_aerop <- read_excel('sales_airport.xlsx')
View(vm_aerop)
vm_aerop <- merge(vm_aerop, novo, by.x = 'ITEM_CODE', by.y = 'COD_PRODUTO', all.x=TRUE)
vm_aerop$mes_entrada <- NULL
vm_aerop$meses[is.na(vm_aerop$meses)] <- 12
vm_aerop$SALES <- round(ifelse(as.numeric(vm_aerop$SALES) < 0, 0, as.numeric(vm_aerop$SALES) )/vm_aerop$meses, 2)
vm_aerop <- rename(vm_aerop, VM = SALES)
##
saldo_aerop <- merge(saldo_aerop, vm_aerop, by = c('ITEM_CODE', 'AIRPORT'), all.x=TRUE)
saldo_aerop$VM[is.na(saldo_aerop$VM)] <- 0
saldo_aerop$meses <- NULL
saldo_aerop$Perm <- round(saldo_aerop$Saldo / saldo_aerop$VM, 2)
saldo_aerop$Risco <- 0
saldo_aerop <- pivot_wider(saldo_aerop, names_from = c('AIRPORT'), values_from = c('Saldo','VM','Perm','Risco'))
saldo_aerop <- rename(saldo_aerop, Artigo = ITEM_CODE)


# ordena colunas por aeroporto
det <- c("Saldo_", "VM_","Perm_", "Risco_")
colunas <- c("Artigo")
for (a in aeroportos) {
  for (d in det) {
    colunas <- append(colunas, paste(d,a,sep=""))
  }
}
saldo_aerop <- saldo_aerop[ , colunas ]


#
View(base)
View(saldo_aerop)
View(vm_aerop)
rm(vm_aerop)


base$`Fec.Inic Permanencia` <- as.Date(as.numeric(base$`Fec.Inic Permanencia`), origin = '1899-12-30')
base$`Saldo disponible calculado` <- as.numeric(base$`Saldo disponible calculado`)
base$`Valor Total disponible` <- as.numeric(base$`Valor Total disponible`)
base$Empresa <- case_when(base$Deposito == 'DA17625216' ~ 'DLF', base$Deposito == 'DA27197888' ~ 'DDB')
base$`Fecha proceso` <- NULL
base$Deposito <- NULL
base$Classificacao <- ifelse(base$Tester == "N", "FOR SALE", "TTG")
base$mes_inicio <- as.numeric(format(base$`Fec.Inic Permanencia`, format = "%Y%m"))


# Base consolidada aging
aging <- base[ , -c(1,2,3,4)]
aging <- aggregate(cbind(aging$`Saldo disponible calculado`, aging$`Valor Total disponible`), list(aging$Empresa, aging$mes_inicio, aging$Categoría, aging$Marca, aging$Articulo, aging$Descripción, aging$Classificacao), FUN = sum)
names(aging) <- c("Empresa", "Mes_inicio", "Categoria", "Marca", "Artigo","Descricao", "Classificacao", "Saldo_Vencer", "Valor_Vencer")

View(aging)

colSums(is.na(aging))


# Recupera venda 12 últimos meses
vm <- read.csv2(file = str_replace("L12_BD&BR_Sales&DamageQty_Aging_MESREF.csv", "MESREF", mes_ref))
vm <- vm[ , c(1,2,8,9)]
vm <- merge(vm, novo, by.x = 'LOCAL_ITEM_CODE', by.y = 'COD_PRODUTO', all.x=TRUE)
vm$mes_entrada <- NULL
vm$meses[is.na(vm$meses)] <- 12
vm$`VMedia Mensal` <- round(ifelse(as.numeric(vm$SALES_QTY) < 0, 0, as.numeric(vm$SALES_QTY) )/vm$meses, 2)
vm$`DESMedia Mensal` <- round(ifelse(as.numeric(vm$DAMAGE_QTY) < 0, 0, as.numeric(vm$DAMAGE_QTY)) /vm$meses, 4)
vm$BU_CODE <- ifelse(vm$BU_CODE == "BR", "DLF", "DDB")
vm <- vm[, c(1,2,6,7)]
names(vm) <- c("Artigo", "Empresa", "VMedia Mensal","DMedia Mensal")
View(vm)
rm(novo)

# Incluindo venda média no DF principal
aging_full <- merge(aging, vm, by = c("Empresa", "Artigo"), all.x = TRUE)
colSums(is.na(aging_full))
aging_full$`VMedia Mensal`[is.na(aging_full$`VMedia Mensal`)] <- 0
aging_full$`DMedia Mensal`[is.na(aging_full$`DMedia Mensal`)] <- 0
View(aging_full)

# Definindo classificação do artigo
aging_full$Descricao <- toupper(aging_full$Descricao)
aging_full$Classificacao[aging_full$Classificacao == "TTG" & grepl(paste(c("BRINDE", "AMOSTRA", "GWP", "BRD ", "GIFT", "BOLSAS TAG EUROPEAS"), collapse = "|"), aging_full$Descricao) ] <- "GWP"
aging_full$Classificacao[aging_full$Classificacao == "TTG" & aging_full$`VMedia Mensal` > 0] <- "GWP"
#aging$Classificacao[aging$Classificacao == "TTG" & grepl(paste(c("PROVADOR", "TESTER", "DEGUSTACAO", "TASTING", "DEGUSTAÇÃO"), collapse = "|"), aging$Descricao) ] <- "TESTER"
aging_full$Classificacao[aging_full$Classificacao == "TTG"] <- "TESTER"
unique(aging_full$Classificacao)

# Para TESTERS a venda média é considerada a média de destruições
aging_full$`VMedia Mensal`[aging_full$Classificacao == "TESTER"] <- aging_full$`DMedia Mensal`[aging_full$Classificacao == "TESTER"]
aging_full$`DMedia Mensal` <- NULL
aging_full$mes_vencimento <- aging_full$Mes_inicio + 500
aging_full <- aging_full[order(aging_full$Empresa, aging_full$Artigo, aging_full$mes_vencimento), ]
aging_full$dt_mes_vencimento <- as.Date(paste(as.character(aging_full$mes_vencimento), "01") , format = "%Y%m%d")

aging_full$meses_para_vencer <- as.numeric(trunc((aging_full$dt_mes_vencimento - dt_base) / 30))
aging_full$consumo_acumulado <- trunc(aging_full$meses_para_vencer * aging_full$`VMedia Mensal`)

aging_full$saldo_acumulado <- ave(aging_full$Saldo_Vencer, paste(aging_full$Empresa, aging_full$Artigo) , FUN = cumsum)
aging_full$perda_vencimento <- aging_full$saldo_acumulado - aging_full$consumo_acumulado

#descontando a perda
aging_full$Saldo_acumulado_menos_perda <- aging_full$saldo_acumulado - ifelse(paste(aging_full$Empresa, aging_full$Artigo) == lag(paste(aging_full$Empresa, aging_full$Artigo)), lag(ifelse(aging_full$perda_vencimento > 0, aging_full$perda_vencimento, 0)), 0)
aging_full$Saldo_acumulado_menos_perda[is.na(aging_full$Saldo_acumulado_menos_perda)] <- aging_full$saldo_acumulado[is.na(aging_full$Saldo_acumulado_menos_perda)]
aging_full$perda_vencimento <- aging_full$Saldo_acumulado_menos_perda - aging_full$consumo_acumulado

aging_full$perda_vencimento <- ifelse(aging_full$perda_vencimento > 0, aging_full$perda_vencimento, 0)
#aging_full$saldo_sem_perda <- aging_full$Saldo_vencimento - aging_full$perda_vencimento

aging_full$risco <- ifelse(aging_full$perda_vencimento > 0, 'SIM', 'NAO')
aging_full$vl_risco <-ifelse(aging_full$perda_vencimento > 0, aging_full$perda_vencimento * (aging_full$Valor_Vencer/aging_full$Saldo_Vencer) , 0) 

sum(aging_full$vl_risco)
sum(aging_full$Valor_Vencer)

#Totais por valor
aging_total <- aggregate( aging_full$vl_risco, list(aging_full$Empresa, aging_full$Classificacao, aging_full$mes_vencimento), FUN = sum)
names(aging_total) = c("Empresa", "Classificacao", "Vencimento", "Valor Risco")
#aging_total$perc_risco <- round(aging_total$`Valor Risco` / aging_total$`Valor Total` * 100, 2)
View(aging_total)
aging_total$`Valor Risco` <- round(aging_total$`Valor Risco`,0)


# Totalizadores por classificacao
temp <- pivot_wider(aging_total, names_from = c("Empresa"), values_from = c("Valor Risco"), values_fill = 0)
temp$Total <- temp$DDB + temp$DLF
temp <- pivot_wider(temp, names_from = c("Classificacao"), values_from = c("Total"), values_fill = 0)
temp$DDB <- NULL
temp$DLF <- NULL
temp <- aggregate(temp[-c(1)], list(temp$Vencimento), FUN = sum)
names(temp)[names(temp) == "Group.1"] <- "Vencimento"
temp$TOTAL <- rowSums(temp[ , 2:ncol(temp)])
View(temp)

aging_total <- pivot_wider(aging_total, names_from = c("Empresa","Classificacao"), values_from = c("Valor Risco"), values_fill = 0)

aging_total <- merge(aging_total, temp, by = c("Vencimento"))

# Totalizadores por categoria
total_cat <- aggregate( aging_full$vl_risco, list(aging_full$Categoria, aging_full$mes_vencimento), FUN = sum)
names(total_cat) = c("Categoria", "Vencimento", "Valor Risco")
#aging_total$perc_risco <- round(aging_total$`Valor Risco` / aging_total$`Valor Total` * 100, 2)
View(total_cat)
total_cat$`Valor Risco` <- round(total_cat$`Valor Risco`,0)
total_cat <- pivot_wider(total_cat, names_from = c("Categoria"), values_from = c("Valor Risco"), values_fill = 0)
total_cat$TOTAL <- rowSums(total_cat[ , 2:ncol(total_cat)])


# Acrescentando saldo por aeroporto
aging_full <- merge(aging_full, saldo_aerop, by = c("Artigo"), all.x = TRUE)
aging_full <- aging_full[order(aging_full$Empresa, aging_full$Artigo, aging_full$mes_vencimento), ]
aging_full$dt_mes_vencimento <- NULL

# Calcula Risco aeroportos
aeroportos_ddb =c("BSB","CNF","FOR","GIG","NAT","REC")
principal_ddb = which(aeroportos =="GIG")
principal_dlf = which(aeroportos =="GRU")

for (i in 1:length(aeroportos)) {
  
  aging_full[[i*4+18]] <- ifelse(((aeroportos[i]=="GIG" & aging_full$Empresa == 'DDB') | (aeroportos[i]=="GRU" & aging_full$Empresa == 'DLF')), 
                                   # se aeroporto principal da empresa
                                   pmin(aging_full[[((i-1)*4)+19]], aging_full$saldo_acumulado) - (aging_full[[((i-1)*4)+20]] * aging_full$meses_para_vencer),
                                   # se filial
                                   case_when(aging_full$Empresa == 'DDB' ~ ifelse(((aging_full[[((principal_ddb-1)*4)+20]]*aging_full$meses_para_vencer < aging_full$saldo_acumulado)|(aging_full[[((principal_ddb-1)*4)+19]]<aging_full$saldo_acumulado))&aeroportos[i] %in% aeroportos_ddb,
                                                                                  pmin(aging_full[[((i-1)*4)+19]], aging_full$saldo_acumulado-(pmin(aging_full[[((principal_ddb-1)*4)+20]] * aging_full$meses_para_vencer, aging_full[[((principal_ddb-1)*4)+19]]))) - (aging_full[[((i-1)*4)+20]]* aging_full$meses_para_vencer),0),
                                             aging_full$Empresa == 'DLF' ~ ifelse(((aging_full[[((principal_dlf-1)*4)+20]]*aging_full$meses_para_vencer < aging_full$saldo_acumulado)|(aging_full[[((principal_dlf-1)*4)+19]]<aging_full$saldo_acumulado))& !(aeroportos[i] %in% aeroportos_ddb),
                                                                                  pmin(aging_full[[((i-1)*4)+19]], aging_full$saldo_acumulado-(pmin(aging_full[[((principal_ddb-1)*4)+20]] * aging_full$meses_para_vencer, aging_full[[((principal_dlf-1)*4)+19]]))) - (aging_full[[((i-1)*4)+20]]* aging_full$meses_para_vencer),0)))

  aging_full[ aging_full[i*4+18] < 0, (i*4+18)] <- 0
}

# salvando resultado
write_xlsx(list(aging_full, aging_total, total_cat), str_replace("Aging_DF_MESREF.xlsx", "MESREF", mes_ref))
#


write.csv2(base, str_replace("base_inperebf_MESREF.csv", "MESREF", mes_ref))
write.csv2(aging, str_replace("base_aging_MESREF.csv", "MESREF", mes_ref))
write.csv2(aging_full, str_replace("Aging_Full_MESREF.csv", "MESREF", mes_ref))
write.csv2(aging_total, str_replace("Aging_Totais_MESREF.csv", "MESREF", mes_ref))

