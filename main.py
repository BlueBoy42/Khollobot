import discord
from discord import app_commands
import json
import datetime
from datetime import timedelta
import pytz
from discord.ext import tasks
from ics import Calendar, Event
import io
import os
import csv

with open("data.json", "r") as f:
    data = json.load(f)

with open("config.json") as f:
    config = json.load(f)

with open("Zone-B.ics", 'r') as f:
    zoneB = Calendar(f.read())

bot = discord.Client(intents=discord.Intents.all())
tree = app_commands.CommandTree(bot)

url = "https://cdn.discordapp.com/icons/883070060070064148/c1880648a1ab2805d254c47a14e9053c.png?size=256&amp;aquality=lossless"
groups = []
khôlles = {}
semaine_collometre = {}

no_kholles_embed = discord.Embed(
    title="Aucune khôlle cette semaine",
    description="Tu n'as pas de khôlles prévues pour cette semaine.",
    colour=discord.Colour.green()
)
no_kholles_embed.set_footer(text="MP2I >>>> MPSI")
no_kholles_embed.set_thumbnail(
    url=url)


def semaine_S():
    """Donne le dictionnaire de correspondance sur le colomètre ou None si elle n'y est pas
    Note: Maintenant calculé automatiquement depuis le CSV dans get_kholles()
    """
    # Si semaine_collometre est déjà rempli par get_kholles(), on ne fait rien
    if semaine_collometre:
        return

    holidays = []

    # Année de début de la periode scolaire, à changer chaque année
    year = config["CurrentYear"]
    for event in zoneB.events:
        date = event.begin.datetime.replace(tzinfo=None)
        if ("Vacances" in event.name) and (datetime.datetime(year, 9, 1) <= date < datetime.datetime(year + 1, 8, 25)):
            # La 1ere semaine de chaque vacance (+1 parce que le début c'est le vendredi) (+1 parce que ce module de ### commence l'année à la semaine 0)
            holidays.append(int(event.begin.datetime.strftime('%W'))+2)
            holidays.append(int(event.end.datetime.strftime('%W')))
    # Semaine de début des khôlles, à changer chaque semestre
    week = config["FirstColleWeek"]
    nb = 0
    while nb <= 15:  # Nombre de semaine de khôlles
        if not ((week) in holidays):
            semaine_collometre[nb] = week
            nb += 1
        week += 1
        if week > int(datetime.datetime(year, 12, 31).strftime('%W')):
            week = 1


def semaine_actuelle() -> int:
    """Fonction renvoyant l'index de la semaine de travail (0-15), ou la prochaine s'il n'y a pas cours cette semaine

    >>> semaine_actuelle()
    3
    """
    if not semaine_collometre:
        semaine_S()
    current_iso_week = datetime.date.today().isocalendar()[1]
    for index, iso_week in semaine_collometre.items():
        if iso_week == current_iso_week:
            return index
    # Alors get la semaine qui est sa bande superieure
    future_weeks = [(index, iso_week) for index, iso_week in semaine_collometre.items() 
                    if iso_week > current_iso_week]
    if future_weeks:
        return min(future_weeks, key=lambda x: x[1])[0]
    return 0


day_to_num = {
    "lundi": 0,
    "mardi": 1,
    "mercredi": 2,
    "jeudi": 3,
    "vendredi": 4,
    "samedi": 5,
    "dimanche": 6  # why not ecoute
}


def get_kholles():
    """Charge les khôlles et groupes depuis le CSV unifié"""
    if not os.path.exists("collometre_data.csv"):
        print("Vous devez convertir votre collomètre !")
        exit()

    global groups, khôlles
    groups = []
    khôlles = {}
    
    with open("collometre_data.csv", 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        mode = None
        
        for row in reader:
            if not row or not row[0]:
                continue
            
            # Détecter la section
            if row[0] == '[GROUPES]':
                mode = 'groupes'
                next(reader)  # Skip header
                continue
            elif row[0] == '[KHOLLES]':
                mode = 'kholles'
                next(reader)  # Skip header
                continue
            
            # Lire les groupes
            if mode == 'groupes':
                groups.append({
                    'group_id': int(row[0]),
                    'membres': [row[1], row[2], row[3]] if len(row) >= 4 else []
                })
            
            # Lire les khôlles
            elif mode == 'kholles':
                semaine_kholle = int(row[5])
                semaine_iso = int(row[6])
                
                # Utiliser semaine_kholle pour la clé (S0-S15 ou S16-S31)
                key_semaine = f"S_{semaine_kholle}"
                if len(row[7]) == 1 and not row[7].isdecimal():
                    continue
                if key_semaine not in khôlles:
                    khôlles[key_semaine] = []
                kholle_data = {
                    'matiere': row[0],
                    'colleur': row[1],
                    'jour': row[2],
                    'heure': row[3],
                    'salle': row[4],
                    'semaine': semaine_kholle,  # S0-S15 ou S16-S31
                    'semaine_iso': semaine_iso,  # Semaine ISO réelle
                    'group_id': int(row[7]) if row[7].isdecimal() else int(row[7][:-1]),
                    'user_id' : -ord("a")+ord(row[7][-1:]) if not row[7].isdecimal() else None
                }
                if kholle_data['group_id'] == 0 :
                    continue
                if len(row) > 9 and row[9]:
                    kholle_data['note'] = row[9]
                
                khôlles[key_semaine].append(kholle_data)
    
    # Construire le mapping semaine_collometre automatiquement
    global semaine_collometre
    semaine_collometre = {}
    for key in sorted(khôlles.keys(), key=lambda x: int(x.split('_')[1])):
        semaine_num = int(key.split('_')[1])
        if khôlles[key]:
            semaine_iso = khôlles[key][0]['semaine_iso']
            # Gérer les deux semestres
            if semaine_num >= 16:
                # Semestre 2: S16 = index 0 du semestre 2
                semaine_collometre[semaine_num - 16] = semaine_iso
            else:
                # Semestre 1: S0 = index 0
                semaine_collometre[semaine_num] = semaine_iso
    return groups, khôlles


def kholles_semaines(user_id: int, semaine: int = semaine_actuelle()) -> list:
    """Sends the week's khôlles for a user_id
    
    Args:
        semaine: Index de la semaine (0-15, cherchera S0-S15 ou S16-S31 automatiquement)
    """
    user_data = data["Members"][str(user_id)]
    user_group_id = user_data["group_id"]

    user_khôlles = []
    
    # Chercher d'abord dans le semestre 1 (S0-S15)
    key_s1 = f"S_{semaine}"
    if key_s1 in khôlles:
        for kholle in khôlles[key_s1]:
            if kholle["group_id"] == user_group_id:
                user_khôlles.append(kholle)
    
    # Si pas trouvé, chercher dans le semestre 2 (S16-S31)
    if not user_khôlles:
        key_s2 = f"S_{semaine + 16}"
        if key_s2 in khôlles:
            for kholle in khôlles[key_s2]:
                if kholle["group_id"] == user_group_id or (list(groups[user_group_id]["membres"]).index(user_data["name"])==kholle["user_id"] and kholle["group_id"] == user_group_id):
                    user_khôlles.append(kholle)
    
    user_khôlles = sorted(user_khôlles, key=lambda x: day_to_num.get(x["jour"], 0))
    return user_khôlles

async def gen_kholle(user_id:int, semaine: int = semaine_actuelle(), custom_char:str="", delta_day:int = -1, colour=discord.Colour.purple(), title:str=""):
    """Dynamicly generates user's colles

    Args:
        user_id (int): User id
        semaine (int, optional): Change the generated kholle's week. Defaults to semaine_actuelle().
        custom_char (str, optional): Edit the message sent. Defaults to "".
        delta_day (int, optional): The delta day of kholles that we want to be appearing (For reminders). Defaults to -1.
        colour (discord.Colour, optional) : Embeds color. Default to discord.Colour.purple()
        title (str, optional) : Embeds title.
    Returns:
        Discord Embed 
    """
    today = datetime.date.today().timetuple().tm_wday
    if delta_day == 2 and today in [5,6]: # If were on saturday or sunday, consider next week
        user_khôlles = kholles_semaines(user_id, semaine_actuelle()+1)
        target_day = (today + 2) % 7
    else:
        user_khôlles = kholles_semaines(user_id, semaine)
        target_day = None
    if not user_khôlles:
        return no_kholles_embed
    embed = discord.Embed(
        title="Tes khôlles pour la semaine" if title == "" else title,
        description=f"Salut, {data["Members"][str(user_id)]["name"].split(" ")[1]}, voici les khôlles que tu as {f"pour la S_{semaine} (Semaine {semaine_collometre[semaine]} de l'année)" if not custom_char else custom_char} : ",
        colour=colour
    )
    for kholle in user_khôlles:
        if target_day is not None:
            if day_to_num[kholle["jour"]] != target_day:
                continue
        elif delta_day!=-1:
            if day_to_num[kholle["jour"]] - today != delta_day: # If delta day and day not in specified range
                continue
        kholle_info = ""
        if "Info" in kholle["matiere"]:
            kholle_info = "**\n[Programme de khôlle](https://nussbaumcpge.be/static/MP2I/pgme.pdf)**"
        if 'Français-Philosophie' in kholle["matiere"]:
            kholle["matiere"] = 'Francais-Philosophie'
        if "Maths" in kholle["matiere"]:
            kholle_info = "**\n[Programme de khôlle de maths                  ](https://cahier-de-prepa.fr/mp2i-thiers/docs?rep=331)**"
        if "Physique" in kholle["matiere"]:
            kholle_info = "**\n[Programme de khôlle de physique](https://cahier-de-prepa.fr/mp2i-thiers/docs?rep=329)**"
        field_value = f"```\nLe {kholle['jour']} à {kholle['heure']}.\n"
        if kholle.get('salle'):
            field_value += f"En salle : {kholle['salle']}\n"
        field_value += "```"
        field_value += kholle_info
        
        # Ajouter
        if kholle.get('note'):
            field_value += f"\n NOTE (IMPORTANT) **{kholle['note']}**"
        
        embed.add_field(
            name=f"{kholle['matiere']} avec {kholle['colleur']}",
            value=field_value,
        )    
    if embed.fields == []:
        return 
    return embed


@bot.event
async def on_ready():
    get_kholles()
    semaine_S()
    await send_reminder_saturday()
    await send_reminder_2days_before()
    await send_reminder_sameday()
    print(f'We have logged in as {bot.user}')
    await tree.sync(guild=None)


@tree.command(name="information", description="Quelques infos sur le bot")
async def info(interaction: discord.Interaction):
    embed = discord.Embed(
        title="Informations",
        description="Voici diverses informations sur le bot"
    )
    embed.add_field(name="Vos données", value="Vos données sont stockés dans un fichier qui n'est pas publique, si vous voulez la supression de vos donneés demandez a l'administrateur du programme")
    embed.add_field(name="Le bot", value="Ce bot a été crée pour donner les khôlles de la mp2i de Thiers, il est opensource, son code source est sur https://github.com/mp2i-thiers/Khollobot")
    embed.set_thumbnail(
        url=url)
    embed.set_footer(text="MP2I >>>> MPSI")

    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="connection", description="Relie ton compte discord aux khôlles")
async def connect(interaction: discord.Interaction):
    if str(interaction.user.id) in data["Members"]:
        data["Members"][str(interaction.user.id)] = {}
        with open("data.json", "w") as f:
            json.dump(data, f, indent=4)

    embed = discord.Embed(
        title="Dans quel groupe es-tu ?",
        description="Choisis ton groupe dans la liste ci-dessous.",
        colour=discord.Colour.purple()
    )
    embed.set_footer(text="MP2I >>>> MPSI")
    embed.set_thumbnail(
        url=url)

    await interaction.response.send_message(embed=embed, view=Select_group(), ephemeral=True)


@tree.command(name="mescolles", description="Affiche tes khôlles prévues pour cette semaine")
async def khôlles_cmd(interaction: discord.Interaction):
    member = data["Members"].get(str(interaction.user.id))

    if not member:
        embed = discord.Embed(
            title="Erreur",
            description="Tu n'as pas encore relié ton compte Discord, ou n'as pas fini ta connexion. Utilise la commande /connection.",
            colour=discord.Colour.purple()
        )
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)

        await interaction.response.send_message(embed=embed, ephemeral=True)
        return

    embed = await gen_kholle(user_id = interaction.user.id, semaine=semaine_actuelle())
    if not embed:
        embed = no_kholles_embed
    await interaction.response.send_message(embed=embed, ephemeral=True, view=select_week())

@tree.command(name="calendrier", description="Créer un fichier ICS de tes colles")
async def calendar_cmd(interaction: discord.Interaction):
    member = data["Members"].get(str(interaction.user.id))

    if not member:
        embed = discord.Embed(
            title="Erreur",
            description="Tu n'as pas encore relié ton compte Discord, ou n'as pas fini ta connexion. Utilise la commande /connection.",
            colour=discord.Colour.purple()
        )
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)

        await interaction.response.send_message(embed=embed, ephemeral=True)
        return

    calendrier = Calendar()
    for week in semaine_collometre:
        user_colles = kholles_semaines(interaction.user.id, week)
        if not user_colles:
            continue
        for kholle in user_colles:
            colle = Event()
            colle.name = f"Khôlle de {kholle["matiere"]}"
            colle.description = kholle["colleur"]
            colle.location = kholle["salle"]

            #Calcul de la date de la colle
            if semaine_collometre[kholle["semaine"]] > 33:
                year = config["CurrentYear"]
            else :
                year = config["CurrentYear"] + 1
            year_start = datetime.datetime(year, 1, 1, tzinfo=pytz.timezone('Europe/Paris'))
            #Premier jour de l'année + n° du jour de la colle - n° Jour de l'année + N° de la semaine de la colle -1 (on sait pas pourquoi -1, mais ça marche)
            date = year_start + datetime.timedelta(days=day_to_num[kholle["jour"]] - year_start.weekday(), weeks=semaine_collometre[kholle["semaine"]]-1) 

            if "-" in kholle["heure"]:
                start, end = kholle["heure"].split("-")
                s_h, s_min = map(int, start.split("h"))
                e_h, e_min = map(int, end.split("h"))
            else :
                s_h, s_min = map(int, kholle["heure"].split("h"))
                e_h, e_min = s_h + 1, s_min - 5
            colle.begin = date + timedelta(hours=s_h, minutes=s_min)
            colle.end = date + timedelta(hours=e_h, minutes=e_min)

            calendrier.events.add(colle)

    buffer = io.BytesIO()
    buffer.write(str(calendrier).encode("utf8"))
    buffer.seek(0)
    fichier_ics = discord.File(fp=buffer, filename=f"calendrier_{member["name"].split(" ")[1]}.txt")

    embed = discord.Embed(
        title="Ton Calendrier",
        description="Voici le calendrier de l'ensemble de tes colles. Le colleur est en description de l'événement et la salle en localisation.",
        colour=discord.Colour.blue()
    )
    embed.add_field(
        name="Fichier ICS",
        value="Au format ICS, un calendrier peut être importé sur la plupart des applications de planification.",
        )
    embed.set_footer(text="MP2I >>>> MPSI")
    embed.set_thumbnail(
        url=url)

    await interaction.response.send_message(embed=embed, file=fichier_ics, ephemeral=True)

class select_week(discord.ui.View):

    def __init__(self):
        super().__init__(timeout=None)
        self.semaine = semaine_actuelle()

    @discord.ui.button(label="Semaine precedente", style=discord.ButtonStyle.danger, emoji="⬅️")
    async def second_button_callback(self, interaction, button):
        self.semaine -= 1

        if self.semaine < 0:
            return no_kholles_embed


        embed = await gen_kholle(semaine=self.semaine, user_id = interaction.user.id)
        view = select_week()
        view.semaine = self.semaine
        await interaction.response.edit_message(embed=embed, view=view)

    @discord.ui.button(label="Semaine suivante", style=discord.ButtonStyle.success, emoji="➡️")
    async def next_week_button_callback(self, interaction, button):
        """
        Button handler to show next week khôlles
        """
        self.semaine += 1
        embed = await gen_kholle(semaine=self.semaine, user_id = interaction.user.id)
        view = select_week()
        view.semaine = self.semaine
        await interaction.response.edit_message(embed=embed, view=view)


async def send_reminder_saturday():
    #Send a remainder every saturday for next week khôlles
    if not (datetime.date.today().timetuple().tm_wday == 5):
        return
    
    # If not school week dont send any message (for holidays)
    current_iso_week = datetime.date.today().isocalendar()[1]
    is_school_week = any(iso_week == current_iso_week for iso_week in semaine_collometre.values())
    if not is_school_week:
        return
    
    for member in data["Members"]:
        if data["Members"][member]["reminder"] != "True":
            continue
        user = await bot.fetch_user(member)

        embed = await gen_kholle(user_id = member, semaine=semaine_actuelle()+1,title="Tes kholles pour la semaine prochaine")

        # To send dms, the app needs to be a bot, not just an app.
        if not embed:
            continue
        await user.send(embed=embed)


async def send_reminder_2days_before():
    # If not school week dont send any message (for holidays)
    current_iso_week = datetime.date.today().isocalendar()[1]
    is_school_week = any(iso_week == current_iso_week for iso_week in semaine_collometre.values())
    if not is_school_week:
        return
    
    for member in data["Members"]:
        if data["Members"][member]["reminder"] != "True":
            continue
        user = await bot.fetch_user(member)
        embed = await gen_kholle(user_id = member, semaine=semaine_actuelle(),colour=discord.Colour.red(), custom_char="pour après demain, prépare la bien ! ", delta_day=2, title="Ta kholle pour après demain")
        if not embed:
            continue
        # To send dms, the app needs to be a bot, not just an app.
        await user.send(embed=embed)

async def send_reminder_sameday():
    # If not school week dont send any message (for holidays)s
    current_iso_week = datetime.date.today().isocalendar()[1]
    is_school_week = any(iso_week == current_iso_week for iso_week in semaine_collometre.values())
    if not is_school_week:
        return
    
    for member in data["Members"]:
        if data["Members"][member]["reminder"] != "True":
            continue
        user = await bot.fetch_user(member)
        embed = await gen_kholle(user_id = member, semaine=semaine_actuelle(),colour=discord.Colour.green(), custom_char="pour aujourd'hui, bonne chance ! ", delta_day=0, title="Ta kholle pour aujourd'hui")
        if not embed:
            continue
        # To send dms, the app needs to be a bot, not just an app.
        await user.send(embed=embed)

class Select_group(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)
        self.add_item(SelectGroupDropdown())


class SelectGroupDropdown(discord.ui.Select):
    def __init__(self):
        options = [
            discord.SelectOption(
                label=f"Groupe {group['group_id']} : {', '.join(group['membres'])}",
                value=str(group["group_id"])
            )
            for group in groups
        ]
        super().__init__(
            placeholder="Choisis ton groupe dans la liste",
            min_values=1,
            max_values=1,
            options=options,
            custom_id="select_group"
        )

    async def callback(self, interaction: discord.Interaction):
        group_id = int(self.values[0])
        selected_group = next(
            (g for g in groups if g["group_id"] == group_id), None)

        embed = discord.Embed(
            title="Qui es-tu ?",
            description=f"Tu es dans le groupe {group_id}. Choisis ton nom dans la liste ci-dessous.",
            colour=discord.Colour.purple()
        )
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)

        await interaction.response.edit_message(embed=embed, view=Select_member(selected_group))


class Select_member(discord.ui.View):
    def __init__(self, group):
        super().__init__(timeout=None)
        self.add_item(SelectMemberDropdown(group))


class SelectMemberDropdown(discord.ui.Select):
    def __init__(self, group):
        options = [
            discord.SelectOption(
                label=member,
                value=member
            ) for member in group["membres"]
        ]
        super().__init__(
            placeholder="Choisis ton nom",
            min_values=1,
            max_values=1,
            options=options,
            custom_id="select_member"
        )
        self.group = group

    async def callback(self, interaction: discord.Interaction):
        member = self.values[0]
        embed = discord.Embed(
            title="C'est noté !",
            description=f"Tu es donc {member}, membre du groupe {self.group['group_id']} !",
            colour=discord.Colour.purple()
        )
        data["Members"][str(interaction.user.id)] = {
            "name": member,
            "group_id": self.group["group_id"]
        }
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)

        with open("data.json", "w") as f:
            json.dump(data, f, indent=4)
        await interaction.response.edit_message(embed=embed, view=ReminderChoiceView(interaction.user.id))


class ReminderChoiceView(discord.ui.View):
    def __init__(self, user_id):
        super().__init__(timeout=None)
        self.user_id = user_id

    @discord.ui.select(
        placeholder="Souhaites-tu recevoir un rappel de ta khôlle ?",
        options=[
            discord.SelectOption(label="Oui", value="True"),
            discord.SelectOption(label="Non", value="False")
        ],
        custom_id="reminder_choice"
    )
    async def select_callback(self, interaction, select):
        choice = select.values[0]
        data["Members"][str(self.user_id)]["reminder"] = choice
        with open("data.json", "w") as f:
            json.dump(data, f, indent=4)
        embed = discord.Embed(
            title="Préférence enregistrée",
            description="Tu recevras un rappel avant ta khôlle." if choice == "True" else "Tu ne recevras pas de rappel avant ta khôlle.",
            colour=discord.Colour.purple()
        )
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)
        await interaction.response.edit_message(embed=embed, view=None)


bot.run(open("token.txt").read().strip())
