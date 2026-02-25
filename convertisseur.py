#!/usr/bin/env python

import pandas as pd
import datetime
import sys
import json
from ics import Calendar

with open("config.json") as f:
    config = json.load(f)

with open("Zone-B.ics", 'r') as f:
    zoneB = Calendar(f.read())


groups = []
khôlles = {}
semaine_collometre = {}


def semaine_S():
    """Donne le dictionnaire de correspondance sur le collomètre ou None si elle n'y est pas"""
    holidays = []
    
    year = config["CurrentYear"]
    for event in zoneB.events:
        date = event.begin.datetime.replace(tzinfo=None)
        if ("Vacances" in event.name) and (datetime.datetime(year, 9, 1) <= date < datetime.datetime(year + 1, 8, 25)):
            holidays.append(int(event.begin.datetime.strftime('%W'))+2)
            holidays.append(int(event.end.datetime.strftime('%W')))
    
    week = config["FirstColleWeek"]
    nb = 0
    while nb <= 15:
        if not ((week) in holidays):
            semaine_collometre[nb] = week
            nb += 1
        week += 1
        if week > int(datetime.datetime(year, 12, 31).strftime('%W')):
            week = 1


def detect_semester(df):
    """Détecte le semestre depuis les colonnes"""
    date_cols = [col for col in df.columns if isinstance(col, (datetime.datetime, pd.Timestamp))]
    
    if date_cols:
        first_date = min(date_cols)
        return 2 if first_date.month >= 2 else 1
    
    s_cols = [col for col in df.columns if isinstance(col, str) and col.startswith('S')]
    if s_cols:
        return 1
    
    return 1


def get_kholles_format1(filepath):
    """Convertit le format 1 (Collomètre + Goupes avec S0-S15)"""
    df1 = pd.read_excel(filepath, sheet_name='Collomètre')
    df2 = pd.read_excel(filepath, sheet_name='Goupes')
    
    semester = detect_semester(df1)
    offset = 0 if semester == 1 else 16
    
    data_khôlles = df1.to_dict(orient="records")
    data_groups = df2.to_dict(orient="records")
    
    current_matiere = None
    for row in data_khôlles:
        if pd.notna(row['Matière']) and pd.isna(row['Colleur']):
            current_matiere = row['Matière']
            continue
        
        if pd.notna(row['Colleur']) and current_matiere:
            colleur = row['Colleur']
            jour = row['Jour'] if pd.notna(row['Jour']) else ''
            heure = row['Heure'] if pd.notna(row['Heure']) else ''
            salle = row['Salle'] if pd.notna(row['Salle']) else ''
            
            for semaine in range(16):
                col_name = f'S{semaine}'
                if col_name in row and pd.notna(row[col_name]):
                    group_id = row[col_name]
                    
                    if group_id in ["p", "P", "I", "i"]:
                        continue
                    
                    group_id = int(group_id) if isinstance(group_id, (int, float)) else group_id
                    semaine_kholle = semaine + offset
                    semaine_iso = semaine_collometre.get(semaine, semaine + config["FirstColleWeek"])
                    
                    key_semaine = f"S_{semaine_kholle}"
                    if key_semaine not in khôlles:
                        khôlles[key_semaine] = []
                    
                    khôlles[key_semaine].append({
                        "group_id": group_id,
                        "matiere": current_matiere,
                        "colleur": colleur,
                        "jour": jour,
                        "heure": heure,
                        "semaine": semaine_kholle,
                        "semaine_iso": semaine_iso,
                        "salle": salle,
                        "note": ''
                    })
    
    # Extraction des groupes
    for row in data_groups[2:]:
        if pd.notna(row['Unnamed: 0']):
            group_a = {
                "group_id": int(row['Unnamed: 0']),
                "eleve1": row['Unnamed: 1'] if pd.notna(row['Unnamed: 1']) else '',
                "eleve2": row['Unnamed: 2'] if pd.notna(row['Unnamed: 2']) else '',
                "eleve3": row['Unnamed: 3'] if pd.notna(row['Unnamed: 3']) else ''
            }
            groups.append(group_a)
        
        if pd.notna(row['Unnamed: 4']):
            group_b = {
                "group_id": int(row['Unnamed: 4']),
                "eleve1": row['Unnamed: 5'] if pd.notna(row['Unnamed: 5']) else '',
                "eleve2": row['Unnamed: 6'] if pd.notna(row['Unnamed: 6']) else '',
                "eleve3": row['Unnamed: 7'] if pd.notna(row['Unnamed: 7']) else ''
            }
            groups.append(group_b)
    
    return groups, khôlles


def get_kholles_format2(filepath):
    """Convertit le format 2 (Semaines + Groupes avec dates)"""
    df1 = pd.read_excel(filepath, sheet_name='Semaines')
    df2 = pd.read_excel(filepath, sheet_name='Groupes')
    
    semester = detect_semester(df1)
    offset = 0 if semester == 1 else 16
    
    date_cols = [(idx, col) for idx, col in enumerate(df1.columns) 
                 if isinstance(col, (datetime.datetime, pd.Timestamp))]
    
    data_khôlles = df1.to_dict(orient="records")
    data_groups = df2.to_dict(orient="records")
    
    current_matiere = None
    
    for row in data_khôlles:
        if pd.notna(row['Matière']) and pd.isna(row['Colleur']):
            current_matiere = row['Matière']
            continue
        
        if pd.notna(row['Colleur']) and current_matiere:
            colleur = row['Colleur']
            jour = row['Jour'] if pd.notna(row['Jour']) else ''
            heure = row['Heure'] if pd.notna(row['Heure']) else ''
            salle = row['Salle'] if pd.notna(row['Salle']) else ''
            
            for s_idx, (col_idx, date_col) in enumerate(date_cols):
                col_name = df1.columns[col_idx]
                if pd.notna(row[col_name]):
                    group_id = row[col_name]
                    
                    if group_id in ["p", "P", "I", "i"]:
                        continue
                    
                    group_id = int(group_id) if isinstance(group_id, (int, float)) else group_id
                    iso_week = date_col.isocalendar()[1]
                    semaine_kholle = s_idx + offset
                    
                    key_semaine = f"S_{semaine_kholle}"
                    if key_semaine not in khôlles:
                        khôlles[key_semaine] = []
                    
                    khôlles[key_semaine].append({
                        "group_id": group_id,
                        "matiere": current_matiere,
                        "colleur": colleur,
                        "jour": jour,
                        "heure": heure,
                        "semaine": semaine_kholle,
                        "semaine_iso": iso_week,
                        "salle": salle,
                        "note": ''
                    })
    
    # Extraction des groupes
    for row in data_groups:
        if pd.notna(row.get('groupe')) != None:
            groupe = {
                "group_id": int(row['groupe']),
                "eleve1": row['eleve1'] if pd.notna(row.get('eleve1')) != None else '',
                "eleve2": row['eleve2'] if pd.notna(row.get('eleve2')) != None else '',
                "eleve3": row['eleve3'] if pd.notna(row.get('eleve3')) != None else ''
            }
            groups.append(groupe)
    
    return groups, khôlles


def detect_format(filepath):
    """Détecte le format du collomètre"""
    xl = pd.ExcelFile(filepath)
    sheet_names = xl.sheet_names
    
    if 'Collomètre' in sheet_names:
        return 'format1'
    # Yes it is hardcodded, but i'll add any other format
    if 'Semaines' in sheet_names:
        return 'format2'
    
    return None


def save_csv(groups, khôlles, output_file):
    """Sauvegarde dans un fichier CSV unifié"""
    with open(output_file, 'w', encoding='utf-8') as f:
        # Section GROUPES
        f.write('[GROUPES]\n')
        f.write('groupe_id,eleve1,eleve2,eleve3\n')
        for groupe in groups:
            f.write(f"{groupe['group_id']},{groupe['eleve1']},{groupe['eleve2']},{groupe['eleve3']}\n")
        
        f.write('\n')
        
        # Section KHOLLES
        f.write('[KHOLLES]\n')
        f.write('matiere,colleur,jour,heure,salle,semaine_kholle,semaine_iso,groupe_id,note\n')
        
        # Flatten toutes les khôlles
        all_kholles = []
        for semaine_key in sorted(khôlles.keys()):
            all_kholles.extend(khôlles[semaine_key])
        
        for kholle in all_kholles:
            f.write(f"{kholle['matiere']},{kholle['colleur']},{kholle['jour']},{kholle['heure']},")
            f.write(f"{kholle['salle']},{kholle['semaine']},{kholle['semaine_iso']},")
            f.write(f"{kholle['group_id']},{kholle['note']}\n")


def convert_collometre(input_file):
    """Fonction principale de conversion"""
    semaine_S()

    format_type = detect_format(input_file)
    
    if not format_type:
        raise Exception
    
    if format_type == 'format1':
        groups_data, kholles_data = get_kholles_format1(input_file)
    else:
        groups_data, kholles_data = get_kholles_format2(input_file)
    
    save_csv(groups_data, kholles_data, "collometre_data.csv")
    
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convertisseur.py <fichier_collometre.xlsx> [output.csv]")
        print("\nExemple: python convertisseur.py Collomètre_S1.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    convert_collometre(input_file)