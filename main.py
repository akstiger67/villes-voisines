import requests
import xlsxwriter

voisin = {}
for cp in range(1000, 99999):
    if cp <= 9999:
        cp = '0' + str(cp)
        print(cp)
    else:
        cp = str(cp)
        print(cp)
    
    rayon = '10'
    myurl = 'https://www.villes-voisines.fr/getcp.php?cp=' + cp + '&rayon=' + rayon
    r = requests.get(myurl)
    if(r.json()) != None:
        a = r.json()


        lst = []

        for i in a:
            if type(i) is str:
                b = str(i)
                # print(a[b]['nom_commune'])
                lst.append([a[b]['nom_commune'], a[b]['code_postal'], a[b]['distance']])
            else:
                lst.append([i['nom_commune'], i['code_postal'], i['distance']]) # we need to check if i is a str or a dict, because the API sends different kind of responses. For example with postcode 01130 you get a key value str:obj and with 01140 you get an object of objects...
        
            y = 10 #taille souhaitée
            for i in range(0, len(lst) - y):
                lst.pop()
        print(lst)
        voisin[cp] = lst
        

print(voisin)      

# création d'un tableau excel vide et remplissage avec les données du tableau
workbook = xlsxwriter.Workbook('communes.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 1

for key in voisin:
    worksheet.write(row, 0, key)
    for item in voisin[key]:
        for i in range(0,len(item)):
            worksheet.write(row, col, item[i])
            col += 1
    row +=1
    col = 1    

workbook.close()



