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
global groups
global kholles
global holidays_iso
groups = {"MP2I":[], "MPI":[]}
kholles = {"MP2I":{},"MPI":{}}
holidays_iso = []
reminders_sent_on = None

no_kholles_embed = discord.Embed(
    title="Aucune khôlle cette semaine",
    description="Tu n'as pas de khôlles prévues pour cette semaine. Profites'en pour bosser",
    colour=discord.Colour.green()
)
no_kholles_embed.set_footer(text="MP2I >>>> MPSI")
no_kholles_embed.set_thumbnail(
    url=url)


def holidays_weeks():
    """Donne la liste des iso des semaines pendant lesquelles n'y a pas cours au au grand Lycée Thiers
    """
    if holidays_iso:
        return

    year = config["CurrentYear"]
    for event in zoneB.events:
        if "Vacances" not in event.name:
            continue
        start = event.begin.datetime.replace(tzinfo=None)
        end = event.end.datetime.replace(tzinfo=None)
        if not (datetime.datetime(year, 9, 1) <= start < datetime.datetime(year + 1, 8, 25)):
            continue
        current = start.date()
        while current <= end.date():
            holidays_iso.append(current.isocalendar().week)
            current += datetime.timedelta(days=1)
    holidays_iso[:] = sorted(set(holidays_iso))
    print(f"Semaines de vacances détectées par Zone-B.ics: {holidays_iso}")


def semaine_actuelle() -> int:
    """Renvoie le numéro ISO de la semaine courante."""
    if not holidays_iso:
        holidays_weeks()
    return datetime.date.today().isocalendar()[1]


day_to_num = {
    "lundi": 0,
    "mardi": 1,
    "mercredi": 2,
    "jeudi": 3,
    "vendredi": 4,
    "samedi": 5,
    "dimanche": 6  # why not ecoute
}


def get_kholles(group_class):
    """Charge les kholles et groupes depuis le CSV correspondant à la classe entrée
    group_class = "MPI" ou "MP2I"
    """
    data_file = f"{group_class.lower()}_data.csv"
    if not os.path.exists(data_file):
        data_file = f"{group_class}_data.csv"
    if not os.path.exists(data_file):
        print(f"Vous devez convertir le collomètre de {group_class}!")
        exit()
    
    with open(data_file, 'r', encoding='utf-8') as f:
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
                groups[group_class].append({
                    'group_id': int(row[0]),
                    'membres': 
                    [row[1], row[2], row[3]] if len(row) >= 4 and row[3] != '' else 
                    [row[1], row[2]] if len(row) >= 4 else
                    []
                })
            
            # Lire les khôlles
            elif mode == 'kholles':
                semaine_kholle = int(row[5])
                semaine_iso = int(row[6])
                
                # Utiliser semaine_iso pour la clé
                if len(row[7]) == 1 and not row[7].isdecimal():
                    continue
                if semaine_iso not in kholles[group_class]:
                    kholles[group_class][semaine_iso] = []
                kholle_data = {
                    'matiere': row[0],
                    'colleur': row[1],
                    'jour': row[2],
                    'heure': row[3],
                    'salle': row[4],
                    'semaine': semaine_kholle,  # Celle indiquée sur le collomètre
                    'semaine_iso': semaine_iso,  # Semaine ISO réelle
                    'group_id': int(row[7]) if row[7].isdecimal() else int(row[7][:-2]),
                    'member_letter': row[7][-1].lower() if not row[7].isdecimal() else None,
                    'user_id': ord(row[7][-1].lower()) - ord("a") if not row[7].isdecimal() else None
                }
                if kholle_data['group_id'] == None :
                    continue
                if len(row) > 9 and row[9]:
                    kholle_data['note'] = row[9]
                
                kholles[group_class][semaine_iso].append(kholle_data)
    
    return groups, kholles


def kholles_semaines(user_id: int, semaine: int = semaine_actuelle()) -> list:
    """Renvoie les kholles de la semaine user_id
    
    Args:
        semaine: Index de la semaine en iso
    """
    user_data = data["Members"][str(user_id)]
    user_group_id = user_data["group_id"]
    user_class = user_data["class"]

    current_week = datetime.date.today().isocalendar()[1]
    if semaine == current_week and current_week in holidays_iso:
        return []

    user_kholles = []
    if semaine in kholles[user_class]:
        for kholle in kholles[user_class][semaine]:
            if kholle["group_id"] != user_group_id:
                continue

            if "Français" in kholle["matiere"] and user_class == "MP2I":
                group = next(
                    (group for group in groups[user_class]
                     if group["group_id"] == user_group_id),
                    None
                )
                if group is None:
                    continue
                member_index = next(
                    (index for index, member in enumerate(group["membres"])
                     if member == user_data["name"]),
                    None
                )
                if member_index != kholle.get("user_id"):
                    continue

            user_kholles.append(kholle)
    
    user_kholles = sorted(user_kholles, key=lambda x: day_to_num.get(x["jour"], 0))
    return user_kholles

async def gen_kholle(user_id:int, semaine:int = semaine_actuelle(), custom_char:str="", delta_day:int = -1, colour=discord.Colour.purple(), title:str="", target_date=None):
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
    today = datetime.date.today().weekday()
    if target_date is not None:
        user_kholles = [
            kholle for kholle in kholles_semaines(
                user_id, target_date.isocalendar().week
            )
            if day_to_num.get(kholle["jour"]) == target_date.weekday()
        ]
        target_day = None
    else:
        user_kholles = kholles_semaines(user_id, semaine)
        target_day = (today + 2) % 7 if delta_day == 2 and today >= 5 else None
    if not user_kholles:
        return no_kholles_embed
    week_label = f"(Semaine {semaine} de l'année)" if not custom_char else custom_char
    embed = discord.Embed(
        title="Tes khôlles pour la semaine" if title == "" else title,
        description=f"Salut, {data['Members'][str(user_id)]['name']}, voici les khôlles que tu as {week_label} : ",
        colour=colour
    )
    embed.set_footer(text="MP2I >>>> MPSI")
    for kholle in user_kholles:
        if target_day is not None:
            if day_to_num[kholle["jour"]] != target_day:
                continue
        elif delta_day!=-1:
            if day_to_num[kholle["jour"]] - today != delta_day: # If delta day and day not in specified range
                continue
        kholle_info = ""
        match kholle["matiere"][0]:
            case "I":
                if data["Members"][str(user_id)]["class"] == "MP2I":
                    kholle_info = f"**\n[Programme de khôlle](https://nussbaumcpge.be/static/MP2I/pgme.pdf)**"
                elif data["Members"][str(user_id)]["class"] == "MPI":
                    kholle_info = "**\n[Programme de khôlle](https://cahier-de-prepa.fr/mpi*-thiers/progcolles?Info)**"
            case "M":
                if data["Members"][str(user_id)]["class"] == "MPI":
                    kholle_info = "**\n[Programme de khôlle](https://cahier-de-prepa.fr/mpi*-thiers/progcolles?Math%C3%A9matiques)**"
                elif data["Members"][str(user_id)]["class"] == "MP2I":
                    kholle_info = "**\n[Programme de khôlle](https://cahier-de-prepa.fr/mp2i-thiers/docs?rep=329)**"
            case "P":
                if data["Members"][str(user_id)]["class"] == "MPI":
                    kholle_info = "**\n[Programme de khôlle](https://mchampion.fr/colles.php)**"
                elif data["Members"][str(user_id)]["class"] == "MP2I":
                    kholle_info = f"**\n[Programme de khôlle](https://cahier-de-prepa.fr/mp2i-thiers/docs?rep=329)**"
        
        field_value = f"```\nLe {kholle['jour']} à {kholle['heure']}.\n"
        if kholle.get('salle'):
            field_value += f"En salle : {kholle['salle']}\n"
        field_value += "```"
        field_value += kholle_info
        
        # Ajouter
        if kholle.get('note'):
            field_value += f"\n NOTE (IMPORTANT) **{kholle['note']}**"
        
        if "Français" in kholle.get('matiere') and data["Members"][str(user_id)]["class"] == "MP2I":
            group = next(
                (group for group in groups[data["Members"][str(user_id)]["class"]]
                 if group["group_id"] == kholle["group_id"]),
                None
            )
            kholle_name = f"Français pour ***{data['Members'][str(user_id)]['name']}***"
        else:
            kholle_name = kholle["matiere"]
        embed.add_field(
            name=f"{kholle_name} avec {kholle['colleur']}",
            value=field_value,
        )
    if embed.fields == []:
        return 

    embed.set_thumbnail(
            url=url)
    return embed


@bot.event
async def on_ready():
    get_kholles("MP2I")
    get_kholles("MPI")
    holidays_weeks()
    await send_reminders()
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
        title="Dans quelle classe es tu (sup/spé) ?",
        description="Choisis ta classe dans la liste ci-dessous.",
        colour=discord.Colour.purple()
    )
    embed.set_footer(text="MP2I >>>> MPSI")
    embed.set_thumbnail(
        url=url)

    await interaction.response.send_message(embed=embed, view=Select_class(), ephemeral=True)


@tree.command(name="mescolles", description="Affiche tes kholles prévues pour cette semaine")
async def kholles_cmd(interaction: discord.Interaction):
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

    today = datetime.date.today()
    viewing_next_week = today.weekday() >= 5
    displayed_week = (
        (today + datetime.timedelta(days=7)).isocalendar().week
        if viewing_next_week
        else semaine_actuelle()
    )
    embed = await gen_kholle(
        user_id=interaction.user.id,
        semaine=displayed_week,
        custom_char="pour la semaine prochaine" if viewing_next_week else "",
        title="Tes khôlles pour la semaine prochaine" if viewing_next_week else ""
    )
    if not embed:
        embed = no_kholles_embed
    view = select_week()
    view.semaine = displayed_week
    await interaction.response.send_message(embed=embed, ephemeral=True, view=view)

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
    for week in kholles[member["class"]]:
        user_colles = kholles_semaines(interaction.user.id, week)
        if not user_colles:
            continue
        for kholle in user_colles:
            colle = Event()
            colle.name = f"Kholle de {kholle["matiere"]}"
            colle.description = kholle["colleur"]
            colle.location = kholle["salle"]

            #Calcul de la date de la colle
            if kholle["semaine_iso"] > 33:
                year = config["CurrentYear"]
            else :
                year = config["CurrentYear"] + 1
            year_start = datetime.datetime(year, 1, 1, tzinfo=pytz.timezone('Europe/Paris'))
            #Premier jour de l'année + n° du jour de la colle - n° Jour de l'année + N° de la semaine de la colle -1 (on sait pas pourquoi -1, mais ça marche)
            date = year_start + datetime.timedelta(days=day_to_num[kholle["jour"]] - year_start.weekday(), weeks=kholle["semaine_iso"]-1) 

            #Le système d'heure n'est jamais uniforme ; formes possibles = "12h15-13h10", "12h15 13h10", "12h15", "12h"
            #On s'interesse donc juste à l'heure du début, à laquelle vient s'associer la minute de la sonnerie
            #Les colles de français utiliseront un système différent

            school_slots = {8:0, 9:0, 10:15, 11:15, 12:15, 13:15, 14:15, 15:15, 16:20, 17:20, 18:20} 
            #Dictionnaire qui à chaque heure, associe la minute de sonnerie de début d'heure 
            #(trouvable sur l'emploi du temps)
            #C'est pas exacte non plus, soyez en avance !

            if kholle["matiere"] != "Français" :
                start_h = int(kholle["heure"][:2])
                if start_h in school_slots.keys():
                    start_min = school_slots[start_h]
                else:
                    start_min = 0
                end_h, end_min = start_h + 1, start_min - 5 
                #55 minutes par "créneaux de kholle"
            
            else:
                if "-" in kholle["heure"]: #cas 1 forme "12h15-13h10" (1er semestre)
                    start, end = kholle["heure"].split("-")
                    start_h, start_min = map(int, start.split("h"))
                    end_h, end_min = map(int, start.split("h"))
                else: #cas 2 forme "12h15" (2eme semestre)
                    start_h, start_min = map(int, kholle["heure"].split("h"))
                    end_h, end_min = start_h + 0, start_min + 30
            
            colle.begin = date + timedelta(hours=start_h, minutes=start_min)
            colle.end = date + timedelta(hours=end_h, minutes=end_min)

            calendrier.events.add(colle)

    buffer = io.BytesIO()
    buffer.write(str(calendrier).encode("utf8"))
    buffer.seek(0)
    fichier_ics = discord.File(fp=buffer, filename=f"calendrier_{member["name"]}.txt")

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
        Button handler to show next week kholles
        """
        self.semaine += 1
        embed = await gen_kholle(semaine=self.semaine, user_id = interaction.user.id)
        view = select_week()
        view.semaine = self.semaine
        await interaction.response.edit_message(embed=embed, view=view)


async def send_reminders():
    """Envoie les rappels correspondant au jour du démarrage."""
    global reminders_sent_on
    today = datetime.date.today()
    if reminders_sent_on == today:
        return
    if today.weekday() == 5:
        targets = [(today + datetime.timedelta(days=7), "la semaine prochaine", None)]
    elif today.weekday() == 6:
        targets = [(today + datetime.timedelta(days=2), "après-demain", discord.Colour.red())]
    else:
        targets = [
            (today, "aujourd'hui", discord.Colour.green()),
            (today + datetime.timedelta(days=1), "demain", discord.Colour.orange()),
            (today + datetime.timedelta(days=2), "après-demain", discord.Colour.red()),
        ]
    print(targets)
    for member in data["Members"]:
        print(data["Members"][member]["reminder"], type(data["Members"][member]["reminder"]))
        if data["Members"][member]["reminder"] != "True":
            continue
        user = await bot.fetch_user(member)

        for target_date, label, colour in targets:
            if target_date.isocalendar().week in holidays_iso:
                continue
            if colour is None:
                embed = await gen_kholle(
                    user_id=member,
                    semaine=target_date.isocalendar().week,
                    title="Tes khôlles pour la semaine prochaine",
                )
            else:
                embed = await gen_kholle(
                    user_id=member,
                    colour=colour,
                    custom_char=f"pour {label}, bonne chance ! ",
                    title=f"Ta khôlle pour {label}",
                    target_date=target_date,
                )
                print(
                    f"Rappel {member} {target_date} ({label}): "
                    f"{'envoyé' if embed is not no_kholles_embed else 'aucune colle'}"
                )
            if embed is not no_kholles_embed:
                await user.send(embed=embed)
    reminders_sent_on = today

class Select_class(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)
        self.add_item(SelectClassDropdown())


class SelectClassDropdown(discord.ui.Select):
    def __init__(self):
        options = [
            discord.SelectOption(
                label= sup_spe,
                value= sup_spe
            )
            for sup_spe in ["MP2I","MPI"]
        ]
        super().__init__(
            placeholder="Choisis ta classe dans la liste",
            min_values=1,
            max_values=1,
            options=options,
            custom_id="select_class"
        )

    async def callback(self, interaction: discord.Interaction):
        selected_class = self.values[0]
        
        embed = discord.Embed(
            title="Dans quel groupe es-tu ?",
            description=f"Tu es en {selected_class}. Choisis ton groupe dans la liste ci-dessous.",
            colour=discord.Colour.purple()
        )
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)

        await interaction.response.edit_message(embed=embed, view=Select_group(selected_class))


class Select_group(discord.ui.View):
    def __init__(self,sup_spe):
        super().__init__(timeout=None)
        self.add_item(SelectGroupDropdown(sup_spe))


class SelectGroupDropdown(discord.ui.Select):
    def __init__(self, sup_spe):
        options = [
            discord.SelectOption(
                label=f"Groupe {group['group_id']} : {', '.join(group['membres'])}",
                value=str(group["group_id"])
            )
            for group in groups[sup_spe]
        ]
        super().__init__(
            placeholder="Choisis ton groupe dans la liste",
            min_values=1,
            max_values=1,
            options=options,
            custom_id="select_group"
        )
        self.sup_spe = sup_spe

    async def callback(self, interaction: discord.Interaction):
        group_id = int(self.values[0])
        selected_group = next(
            (g for g in groups[self.sup_spe] if g["group_id"] == group_id), None)

        embed = discord.Embed(
            title="Qui es-tu ?",
            description=f"Tu es dans le groupe {group_id}. Choisis ton nom dans la liste ci-dessous.",
            colour=discord.Colour.purple()
        )
        embed.set_footer(text="MP2I >>>> MPSI")
        embed.set_thumbnail(
            url=url)

        await interaction.response.edit_message(embed=embed, view=Select_member(selected_group, self.sup_spe))


class Select_member(discord.ui.View):
    def __init__(self, group, sup_spe):
        super().__init__(timeout=None)
        self.add_item(SelectMemberDropdown(group, sup_spe))


class SelectMemberDropdown(discord.ui.Select):
    def __init__(self, group, sup_spe):
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
        self.sup_spe = sup_spe

    async def callback(self, interaction: discord.Interaction):
        member = self.values[0]
        embed = discord.Embed(
            title="C'est noté !",
            description=f"Tu es donc {member}, en {self.sup_spe}, membre du groupe {self.group['group_id']} !",
            colour=discord.Colour.purple()
        )
        data["Members"][str(interaction.user.id)] = {
            "name": member,
            "group_id": self.group["group_id"],
            "class": self.sup_spe
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
