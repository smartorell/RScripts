library(readxl)
library(writexl)
library(readr)
library(stringr)

# gera lista de arquivos que se quer trabalhar
lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = "Inventory.*.xlsx")
lista_arquivos

rm(pos_mensal)
for (arq in lista_arquivos) {
  # lista abas do arquivo em questão
  abas = excel_sheets(arq)
  linha_inicio = 9
  for (aba in abas) {
    temp <- read_excel(path = arq, sheet = aba, col_types = "text", range =  cell_rows(c(linha_inicio, NA))) 
    linhas_aba = nrow(temp)
    # remove linhas cuja qtde em estoqueestá vazia
    temp <- temp[!is.na(temp$`On Consignment Units`), ]
    # hremove colunas que não serão usadas
    temp <- temp[ , -c(18:47)]
    temp <- temp[ , -c(4,5,9,10,22,23, 24)]
    temp$`Unidade Negocio` = substr(temp$`Shop Code`, 1, 2)
    if (exists("pos_mensal")) {
      pos_mensal <- rbind(pos_mensal, temp)
    }  
    else {
      pos_mensal <- temp
    }
    linha_inicio = 8
  }
}  

View(pos_mensal)
pos_mensal[ , 14:17] <- sapply(pos_mensal[ , 14:17], as.numeric)
pos_mensal$Qtde_Total = pos_mensal$`On Hand Units` + pos_mensal$`On Consignment Units`
pos_mensal <- pos_mensal[pos_mensal$Qtde_Total !=  0, ]

# salvando resultado final
nome_arq_saida = paste("Estoque Vegga DP-", str_replace(pos_mensal$DATE[1], "/","-"), sep = "")
nome_arq_saida = paste(nome_arq_saida, ".csv", sep = "")
nome_arq_saida

# salvando resultado
write_xlsx(pos_mensal, str_replace(nome_arq_saida, ".csv", ".xlsx"))
write_csv2(pos_mensal, nome_arq_saida)
#

# consolidando por artigo
saldo_artigo <- aggregate(cbind(pos_mensal$`On Hand Units`, pos_mensal$`On Consignment Units`, pos_mensal$`On Hand Value`, pos_mensal$`On Consignment Value`, pos_mensal$Qtde_Total), list(pos_mensal$`Shop Code`, pos_mensal$`Gamma Product Code`), FUN = sum)
names(saldo_artigo) <- c("Centro", "ArtigoGamma", "Qtde Propria", "Qtde Consignada", "Val Proprio", "Val Consignado", "Qtde Total")

write_csv2(saldo_artigo, 'saldo_comercial_DP_artigo.csv')


# consultas no resultado final
sum(pos_mensal$Qtde_Total)
sum(pos_mensal$Qtde_Total[pos_mensal$Tester == "N"])







