import Diccionari

key_list = list(Diccionari.classificació.keys())
print(key_list)

num_files = len(key_list)

for i in range(2, num_files + 2):  ## ------- Crear la taula a partir del diccionari
    for j in key_list:
        ws2.cell(row=i, column=j).value = j

