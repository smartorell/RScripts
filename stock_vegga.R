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
    temp <- temp[ , -c(18:50)]
    temp <- temp[ , -c(4,5,9,10,19,23)]
    temp$`Unidade Negocio` = substr(temp$`Shop Code`, 1, 2)
    # filtra linhas com estoque diferente de zero
    temp <- temp[temp$`On Consignment Units` !=  0, ]
    linhas_limpas = nrow(temp)
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

# salvando resultado final
nome_arq_saida = paste("Estoque Vegga DF-", str_replace(pos_mensal$DATE[1], "/","-"), sep = "")
nome_arq_saida = paste(nome_arq_saida, ".csv", sep = "")
nome_arq_saida

# salvando resultado
write_xlsx(pos_mensal, str_replace(nome_arq_saida, ".csv", ".xlsx"))
write_csv2(pos_mensal, nome_arq_saida)
#

# consolidando por artigo
saldo_artigo <- aggregate(cbind(pos_mensal$`On Consignment Units`, pos_mensal$`On Consignment Value`, pos_mensal$`Transit Units`, pos_mensal$`Transit Value`), list(pos_mensal$`Gamma Product Code`), FUN = sum)
names(saldo_artigo) <- c("ArtigoGamma", "Qtde", "FOB_tot", "Qtde Transito", "Val Transito")

write_csv2(saldo_artigo, 'saldo_comercial_artigo.csv')


# consultas no resultado final
sum(pos_mensal$`On Consignment Units`)
sum(pos_mensal$`On Consignment Units`[pos_mensal$Tester == "N"])





