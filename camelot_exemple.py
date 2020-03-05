import camelot
import pprint
pp = pprint.PrettyPrinter(indent=4)

# Créer une tableList camelot contenant le ou les tableaux du document passé en argument.
# split_text sur True permet de séparer les textes que camelot regroupe de colone différentes car trop proche les un des autres 
#pages='all' lis tout le document
tables = camelot.read_pdf('102955.pdf', pages='all', split_text=True, line_scale=40) #flavor='stream',,columns=[10

# retourn le nombre de tableau annalysé par camelot
len(tables)

# table[1]# data du premier tableau

x = tables[0].parsing_report #Permet de faire états de la qualité de l'extraction
print(x)

#permet de boucler sur les tableau
# for table in tables:
#     table.data # retourn un liste de liste correspondant aux donnée contenu dans le tableau.



import matplotlib.pyplot as plt
for table in tables:
    x= camelot.plot(table, kind='joint')
    
    pp.pprint(table.data)
    plt.show(x)

# tables[0].to_csv('foo1.csv') # créer et inject un tableau dans un CSV

