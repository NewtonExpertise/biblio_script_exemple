#coding:utf-8
import xlwings as xw
import re

def instance_excel():
    """
    de l'instance Excel aux feuilles des classeurs
    """
    # wb = xw.Book(r'V:\Mathieu\v.xlsx')
    # xw.apps.keys() récupère les instances excel ouvertes sous forme de list.
    instance = xw.apps.keys()

    print(instance)
    # Nous bouclons dessus pour récupérer chaque instance.
    for i in instance:
        # Pour chaque instance nous récupérons le workbook excel.
        for book in xw.apps[i].books:
            print(book)
            print(book.fullname)
            for sheet in book.sheets:
                print(sheet)
                print(sheet.name)
                print(sheet.book)


# # Méthode pour générer un onglet unique
def add_sheet_new_name(wb, nom):
    """
    Génère une feuille excel avec un nom unique
    nb : une feuille excel ne peut contenir que 31 caractères
    """
    nom = nom[:30]
    increment = 0
    sheet = [sheet.name for sheet in wb.sheets]
    if nom in sheet:

        for sheet_name in sheet:
            try:
                name, compteur , _ =  re.split(nom+r'(\d+)',sheet_name)
            except ValueError:
                compteur =0
            if increment< int(compteur):
                increment = int(compteur)

        new_name = nom + str(increment+1)
    else:
        new_name = nom
    new_ws = wb.sheets.add(new_name)
    new_ws.active

    return new_ws

def range_pour_fonction_xl():
    """
    Permet de définir des plages de cellules/ colonns
    """

    ws = xw.sheets.active
    wb = ws.book
    xw.Range('A1:E10').name = 'bonjour'

def lecture_sheet_xl():
    """
    Méthode pour lire un fichier excel
    """
    # retourn un dictionnaire E clé F valeur
    sheet.range('E1:F500').options(dict, numbers=int).value

    # retourn la valeur d'une cellule
    sheet.range('A1').value


def set_validations_donnee():


    bases_name = ["Bonjour","Je","Suis","Une","Liste"]
    str_bases_name = ';'.join(bases_name)
    xw.Range('A1').api.validation.delete() # suppression de l'ancienne liste déroulante
    xw.Range('A1').clear() # suppression de la valeur de la cellule
    xw.Range('A1').api.validation.add(3,2,3,str_bases_name) #Création de la liste déroulante

def dissimule_feuille_excel():

    wb = xw.Book('_PandL.xlsx')

    wb.sheets["Mcdo 1"].api.Visible = False

def formule_de_calcule():
    """
    insérer une formule sous excel
    """
    ws = xw.sheets.active
    ws.range('G1').value = '=A2+5'

def soustoto():
    """
    Créer une vue sous toto
    """
    ws = xw.sheets.active
    nbligne = ws.cells(ws.api.rows.count, "A").end(-4162).row
    ws.range("A1", "D"+str(nbligne)).api.subtotal(3, -4157,4, True, False,True)
    ws.autofit()
    ### ws.range("A1", "E145").api.subtotal( (indiquez la colone a regrouper), (méthode de calcule) , (tableau contenant la position des lignes a regrouper) , True, False,True)

def derniere_ligne():
    """
    retourn la dernière ligne du tableau
    """
    ws = xw.sheets.active
    nbligne = ws.cells(ws.api.rows.count, "A").end(-4162).row
    return nbligne

def filtre_Tableau():
    """
    effectue un filtre/trie par ordre croissant sur la colonne A sur un tableau excel de A2 à D fin de ligne.
    """
    sheet = xw.sheets.active
    # wb = ws.book
    nbligne = sheet.cells(sheet.api.rows.count, "A").end(-4162).row
    sheet.range('A2:D'+str(nbligne)).api.Sort(Key1=sheet.range("A:A").api, Order1=1)

def set_filter_header():
    """
    Ajoute un filtre sur les en-tête du tableau
    """
    ws = xw.sheets.active
    nbligne = ws.cells(ws.api.rows.count, "A").end(-4162).row
    ws.range("A1", "D"+str(nbligne)).api.AutoFilter(VisibleDropDown=True)

def pivotTable():
    ws = xw.sheets.active
    wb = ws.book
    ws = wb.sheets['Feuil1']
    try:
        ws_dyna = wb.sheets.add('new_sheet')
    except:
        ws_dyna = wb.sheets['new_sheet']
    # nbligne = ws.cells(ws.api.rows.count, "A").end(-4162).row
    # ws.range("A1", "D"+str(nbligne)).api.PivotTable

    # Création d'un cache pour les TCD
    cache = wb.api.PivotCaches().Create(SourceType= 1 ,SourceData=ws.range("A1:D16").api )

    # Création d'un TCD (cette ligne créer un TCD sur une nouvelle feuille excel. Aucun champ n'est assigné le TCD est vide.)
    ws.api.PivotTableWizard( SourceType=1 , SourceData=ws.range("A1:D16").api , TableDestination=ws_dyna.range("F5").api , RowGrand=True, ColumnGrand=True, HasAutoFormat= True, TableName= "TP Mathieu")

    # Assignation des donnée a calculer pour TCD 
    ws_dyna.api.PivotTables("TP Mathieu").AddDataField(Field =  ws_dyna.api.PivotTables("TP Mathieu").PivotFields("vert"), Caption = "somme vert", Function= -4157)
    ws_dyna.api.PivotTables("TP Mathieu").AddDataField(Field =  ws_dyna.api.PivotTables("TP Mathieu").PivotFields("bleu"), Caption = "somme bleu", Function= -4157)
    ws_dyna.api.PivotTables("TP Mathieu").AddDataField(Field =  ws_dyna.api.PivotTables("TP Mathieu").PivotFields("jaune"), Caption = "somme jaune", Function= -4157)
    # Assignation des données a afficher en ligne.
    ws_dyna.api.PivotTables("TP Mathieu").AddFields(RowFields = "compte")
    # Set une orientation des données calculer en colonne
    ws_dyna.api.PivotTables("TP Mathieu").DataPivotField.Orientation = 2
    ws_dyna.api.PivotTables("TP Mathieu").DataPivotField.Position= 1
    # Set un style au tabeau croisé dynamique
    ws_dyna.api.PivotTables("TP Mathieu").TableStyle2 = "PivotStyleMedium9"



    # # # # cache = wb.api.PivotCaches().Add(SourceType=1, SourceData=ws.range("A1:D16"))
    # # # # cache.CreatePivotTable(TableDestination='new_sheet', )
    # ws.PivotTables.Add(cache, ws_dyna, "test_pyvo")
    # .CreatePivotTable(TableDestination =ws_dyna.range("A3") , TableName='test_dyna')
    # ws_dyna.select().range('A3').select()
    # ws_dyna.api.PivotTables('test_dyna').PivotFields("individus")




if __name__ == "__main__":
    print(derniere_ligne())
