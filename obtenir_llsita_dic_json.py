import json
llista_clas = []
with open('classificacio.json', encoding='utf-8') as json_file:
    data = json.load(json_file)
    dic = data["classificacio"]
    for elem in dic:
        for y in data["classificacio"][elem]:
            print(elem, y)
    if "Tesla" in dic:
        print("esta")