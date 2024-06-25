?toJSON

i = lista_arquivos[1]

env = read_lines(file = i, skip = 4, n_max = -1)
json_obj = toJSON(env)
json_obj = fromJSON(env)
con$insert(json_obj)

?Sys.setenv

?break

env = read_lines(file = 'CBDA901200107-BDA901-27197888008135-202206220317167.ENV', skip = 4, n_max = -1)
con$insert(env)
print(json_obj)