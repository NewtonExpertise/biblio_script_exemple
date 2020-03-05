from datetime import datetime, timedelta

def end_of_month(dt0):
    """
    Renvoi le dernier jour du mois de la date donnée
    """
    dt1 = dt0.replace(day=1)
    dt2 = dt1 + timedelta(days=32)
    dt3 = dt2.replace(day=1) - timedelta(days=1)
    return dt3

def listing_fin_mois_sur_periode():
    a = datetime(2019, 1, 1, 0, 0)
    b = datetime(2020, 12, 30, 0, 0)

    test_liste=[]
    while debut <= fin :
        test_liste.append(end_of_month(debut))
        debut = end_of_month(debut)+ timedelta(1)

    return test_liste