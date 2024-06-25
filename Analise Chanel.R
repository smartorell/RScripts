library(dplyr)

# Analise artigos Chanel
# cria DF somente com artigos marca Chanel
chanel <- pos_mensal[(pos_mensal$`Global Brand` == "CHANEL" ), ]
chanel <- chanel[!is.na(chanel$`On Consignment Units`), ]
View(chanel)

# consultas resultado chanel
sum(chanel$`On Consignment Units`)
sum(chanel$`On Consignment Units`[chanel$`Unidade Negocio` == "BD"])
sum(chanel$`On Consignment Units`[chanel$Tester == "N" & chanel$`Shop Code` == "BRT072"])
chanel[chanel$Tester == "N" & chanel$`On Consignment Units` < 0, ]

# quantidade por loja
chanel  <- chanel[chanel$Tester == "N" , ]
stock_loja <- aggregate(chanel$`On Consignment Units`, list(chanel$`Shop Code`), FUN = sum)
stock_loja  <- rename(stock_loja, Centro = Group.1)
stock_loja <- rename(stock_loja, Qtde = x)
stock_loja

# lista arquivos a serem pesquisados
artigos_chanel <- read_excel("artigos_chanel.xlsx")
View(artigos_chanel)  

# faz join da posição de estoque chanel com os artigos contados no inventario
artigos_contagem <- merge(chanel, artigos_chanel, by.x = "Product Code", by.y = "ITEM")
sum(artigos_contagem$`On Consignment Units`)
View(artigos_contagem)