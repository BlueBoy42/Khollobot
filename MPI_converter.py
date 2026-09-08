import pandas as pd
import datetime
import sys
import json
from ics import Calendar
import pdfplumber

with open("config.json") as f:
    config = json.load(f)

with open("Zone-B.ics", 'r') as f:
    zoneB = Calendar(f.read())

day = {"Lun":"lundi",
       "Mar":"mardi",
       "Mer":"mercredi",
       "Jeu":"jeudi",
       "Ven":"vendredi"
        }

khôlles = {}
semaine_collometre = {}
groups = []



def semaine_S():
    """Donne le dictionnaire de correspondance sur le collomètre ou None si elle n'y est pas"""
    year = config["CurrentYear"]
    for s, l in zip(Semaines,Lundi):
        j, m = l.split("/")
        j, m = int(j), int(m)
        if m > 8 :
            date_lundi = datetime.date(year,m,j)
        else :
            date_lundi = datetime.date(year + 1,m,j)
        semaine_collometre[s] = date_lundi.isocalendar()[1]
        

# Ouvrir le PDF
with pdfplumber.open("Colloscope.pdf") as pdf:
    # Accéder à la première page
    page = pdf.pages[0]
    
    # Extraire les tableaux
    table = page.extract_tables()[0]

    Semaines = table[3][4:-5]
    Lundi = table[2][4:-5]



def get_kholles():
    Colleur = []
    Jour = []
    Heure = []
    Salle = []
    Matière = []
    mem_colleur = ""
    mem_matière = ""
    for row in table[5::]:
        #les paramètres

        if row[0] != None:
            mem_matière = row[0][::-1]
        if row[1] != None:
            mem_colleur = row[1]
        jour, horaire = row[2].split(" ")
        Jour.append(day[jour])
        Heure.append(horaire[:3])
        Matière.append(mem_matière)
        Colleur.append(mem_colleur)
        Salle.append(row[3]) 

    #les colles
    for col in range(len(Semaines)):
        khôlles[Semaines[col]] = []

        semaine_iso = semaine_collometre[Semaines[col]]
        for row in range(len(Colleur)):
            groupe_id = table[row + 5][col + 4]
            if groupe_id != '':
                khôlles[Semaines[col]].append({
                    "groupe_id": groupe_id,
                    "matiere": Matière[row],
                    "colleur": Colleur[row],
                    "jour": Jour[row],
                    "heure": Heure[row],
                    "semaine": Semaines[col],
                    "semaine_iso": semaine_iso,
                    "salle": Salle[row],
                    "note": ''
                })
    return groups, khôlles


#-------Groupes--------
    
for row in table[5:23]:
    if row[28] != '':
        groupe = {
            "groupe_id" : int(row[28]),
            "eleve1": row[29],
            "eleve2": row[30], 
            "eleve3": row[31],
        }
        groups.append(groupe)


def save_csv(groups, khôlles, output_file):
    """Sauvegarde dans un fichier CSV unifié"""
    with open(output_file, 'w', encoding='utf-8') as f:
        # Section GROUPES
        f.write('[GROUPES]\n')
        f.write('groupe_id,eleve1,eleve2,eleve3\n')
        for groupe in groups:
            f.write(f"{groupe['groupe_id']},{groupe['eleve1']},{groupe['eleve2']},{groupe['eleve3']}\n")
        
        f.write('\n')
        
        # Section KHOLLES
        f.write('[KHOLLES]\n')
        f.write('matiere,colleur,jour,heure,salle,semaine_kholle,semaine_iso,groupe_id,note\n')
        
        # Flatten toutes les khôlles
        all_kholles = []
        for semaine_key in khôlles.keys():
            all_kholles.extend(khôlles[semaine_key])
        
        for kholle in all_kholles:
            f.write(f"{kholle['matiere']},{kholle['colleur']},{kholle['jour']},{kholle['heure']},")
            f.write(f"{kholle['salle']},{kholle['semaine']},{kholle['semaine_iso']},")
            f.write(f"{kholle['groupe_id']},{kholle['note']}\n")


def convert_collometre(input_file):
    """Fonction principale de conversion"""
    semaine_S()

    groups_data, kholles_data = get_kholles()
    
    save_csv(groups_data, kholles_data, "collometre_data.csv")
    
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python MPI_converter.py <fichier_collometre.pdf> [output.csv]")
        print("\nExemple: python MPI_converter.py Colloscope.pdf")
        sys.exit(1)
    
    input_file = sys.argv[1]
    convert_collometre(input_file)
    