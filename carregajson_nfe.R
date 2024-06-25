library(readr)
library(dplyr)
library(mongolite)
library(rjson)

Sys.setenv("VROOM_CONNECTION_SIZE"= 262144*4)

# função para carregar diversos arquivos ENV
carrega_json <- function (p_con, p_padrao) {
  lista_arquivos <- list.files(getwd(), full.names = TRUE, pattern  = p_padrao)
  for (i in lista_arquivos) {
    env = read_lines(file = i, skip = 4, n_max = -1)
    p_con$insert(env)
  }
}

# Cria conexao com MongoDB
con <- mongo(collection = 'nfe', db = 'dufry',
             url = 'mongodb://localhost', verbose = FALSE,
             options = ssl_options())

carrega_json(con, "*.ENV")

con$count()

#Apaga conteudo completo da coleção
con$remove(query = "{}", just_one = FALSE)
