import re

def get_unique_path(path):
    """
    Retourn un nom de fichier unique en fonction des fichiers déjà existant dans le dossier
    exemple : fichier.txt , fichier (1).txt etc
    """
    # si le nom de fichier existe, on en cherche un autre
    while os.path.exists(path):
        # on vire l'extension
        base, ext = os.path.splitext(path)
        try:
           # on extrait le compteur si il existe
           base, counter, _ = re.split(r" \((\d+)\)$", base)
        except ValueError:
           counter = 0
        # on reconstruit le path
        path = "%s (%s)%s" % (base, int(counter) + 1, ext)
    return path