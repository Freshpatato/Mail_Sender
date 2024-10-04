import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Text
from ttkbootstrap import Style
from ttkbootstrap.constants import *
from ttkbootstrap import ttk
import configparser
import webbrowser
from datetime import datetime
import shutil


# --- Variables Globales ---
presets = {}
dota_file = None
filtered_df = None
global df

current_datetime = datetime.now()
formatted_datetime = current_datetime.strftime("%d/%m/%Y %H:%M:%S")

loaded_file_path = None  # Variable pour stocker le chemin du fichier chargé
attachment_file_path = None  # Variable pour stocker le chemin de la pièce jointe
hp_file_path = None #Variable chemin fichier HP


# --- Gestion configuaration ---
config_directory = os.path.join(os.path.expanduser("~"), ".retro_app")
os.makedirs(config_directory, exist_ok=True)
config_file = os.path.join(config_directory, 'config.ini')
log_file = os.path.join(config_directory, 'log.txt')


def read_config():
    config = configparser.ConfigParser()
    config.read(config_file)

    if 'SMTP' in config:
        username_entry.insert(0, config['SMTP'].get('username', ''))
        password_entry.insert(0, config['SMTP'].get('password', ''))
        from_address_entry.insert(0, config['SMTP'].get('from_address', ''))

    if 'Presets' in config:
        for key in config['Presets']:
            try:
                preset_type, subject, content = config['Presets'][key].split('::', 2)
                presets[key] = (preset_type, subject, content)
                preset_listbox.insert(tk.END, f"{key} ({preset_type})")
            except ValueError:
                print(f"Erreur de format pour le preset {key}")

    if 'Attachment' in config:
        global attachment_file_path
        attachment_file_path = config['Attachment'].get('pdf', '')
        attachment_label.config(text=os.path.basename(attachment_file_path))
    
    if 'File_HP' in config:
        global hp_file_path
        hp_file_path = config['File_HP'].get('xlsx', '')
        hp_file_label.config(text=os.path.basename(hp_file_path))
                                              

    preset_combobox['values'] = list(presets.keys())


def write_config():
    config = configparser.ConfigParser()
    if not os.path.exists(config_file):
        with open(config_file, 'w') as f:
            f.write("")
    config.read(config_file)

    if 'SMTP' not in config:
        config['SMTP'] = {
            'username': username_entry.get(),
            'password': password_entry.get(),
            'from_address': from_address_entry.get()
        }
    else:
        config['SMTP']['username'] = username_entry.get()
        config['SMTP']['password'] = password_entry.get()
        config['SMTP']['from_address'] = from_address_entry.get()

    if not config.has_section('Presets'):
        config.add_section('Presets')
    else:
        config.remove_section('Presets')
        config.add_section('Presets')

    for key, value in presets.items():
        config.set('Presets', key, f"{value[0]}::{value[1]}::{value[2]}")

    if 'Attachment' not in config:
        config['Attachment'] = {
            'pdf': str(attachment_file_path)
        }
    else:
        config['Attachment']['pdf'] = str(attachment_file_path)

    if 'File_HP' not in config:
        config['File_HP'] = {
            'xlsx': str(hp_file_path)
        }
    else:
        config['File_HP']['xlsx'] = str(hp_file_path)

    with open(config_file, 'w') as configfile:
        config.write(configfile)



def log_action(action):
    with open(log_file, "a", encoding="utf-8") as log_file_obj:
        log_file_obj.write(action + "\n")


def delete_log_file():
    if os.path.exists(log_file):
        os.remove(log_file)
        messagebox.showinfo("Info", "Log vidé")
        read_file_log()
    else:
        messagebox.showerror("Info", "Log déjà vidé")

def read_version():
    # Chemin du répertoire de configuration où se trouve config.ini et version.txt
    config_directory = os.path.join(os.path.expanduser("~"), ".retro_app")
    version_file_path = os.path.join(config_directory, "version.txt")  

    try:
        with open(version_file_path, 'r') as file:
            version = file.read().strip()
            return version
    except FileNotFoundError:
        return "Version inconnue"  
    except OSError as e:
        return f"Erreur: {e}" 
    

# --- Configuration SMTP ---
def config_smtp():
    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    username = username_entry.get()
    password = password_entry.get()
    from_address = from_address_entry.get()
    
    # Configuration du serveur SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)
    
    return server, from_address


# --- Getion excel Dotation/Remplacement ---
def import_excel_file_endowment():
    global file_path_endowment
    file_path_endowment = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    load_excel_file_endowment(file_path_endowment)



def load_excel_file_endowment(file_path):   
    if not file_path:
        return
    global dota_file
    file_path_endowment = file_path  
    dota_file = os.path.join(config_directory, 'tmp_dotation.xlsx')
    shutil.copy2(file_path_endowment, dota_file)
    path_excel_label.config(text=file_path_endowment.split("/")[-1])
    read_excel_file_endowment(dota_file)


# Lire fichier excel dotation/rempla
def read_excel_file_endowment(file_path):
    global filtered_df  
    if not file_path:
        return
    df = pd.read_excel(file_path, sheet_name='Agence')
    df.columns = df.iloc[0]
    df = df[1:]
    selected_columns = df[[ 
        'Demandeur', 'NOM,Prénom', 'Mail', 'Ancien DST', 'Ville', 'DST', 'Ref exp', 'Service expe étranger', 'Rempla/ Dotation', 'Logiciels']]
    
    # Effacer les anciennes données du Treeview
    for row in treeview.get_children():
        treeview.delete(row)
    
  
    filtered_df = selected_columns[(
        selected_columns['Rempla/ Dotation'].str.contains('Rempla EZV|Dotation EZV', case=False, na=False)) & (
        selected_columns['Logiciels'].str.contains('Mail', case=False, na=False)) ]
   
    # Insérer les données dans le Treeview
    for index, row in filtered_df.iterrows():
        treeview.insert("", "end", values=(
            row['Demandeur'], 
            row['Mail'], 
            row['NOM,Prénom'], 
            row['Ville'], 
            row['Ref exp'], 
            row['DST'], 
            row['Ancien DST']
        ))



# --- Message cloture Doation/Remplacement
def copy_selected_email():
    selected_items = treeview.selection()  # Récupérer les éléments sélectionnés dans le Treeview
    if not selected_items:
        messagebox.showerror("Erreur", "Aucun email sélectionné")
        return

    for item in selected_items:
        values = treeview.item(item, 'values')  # Récupérer les valeurs des colonnes pour la ligne sélectionnée
        if len(values) >= 6:
            demandeur, mail_to, beneficiare, ville, expe, dst = values[:6]
            old_dst = values[6] if len(values) == 7 else ""
            print(old_dst)
            if old_dst.strip().lower() == 'dotation':
                # Si old_dst n'existe pas, message dotation
                text = f"""
                Bonjour,

                Votre PC a été expédié. ({dst})
                Lieu : {ville}

                Cordialement,
                Postes de Travail France
                """
            else:
                if var_dpd.get() == 0:  # Prepa Venissieux
                    text = f"""
                    Bonjour,
                    
                    Le nouveau PC {dst} a été expédié ce jour à votre attention.
                    
                    Concernant la restitution de votre ancien PC {old_dst}, il vous suffit de prendre rendez-vous via notre outil RESERVIO.
                    Pour accéder à l’outil de réservation, merci de prendre connaissance du mail « AVIS EXPEDITION »
                    
                    Cordialement,
                    Postes de Travail France
                    """
                else:  # Préparation à distance
                    text = f"""
                    Bonjour,
                    
                    Le nouveau PC {dst} a été expédié ce jour à votre attention.
                    
                    Concernant la restitution de votre ancien PC {old_dst}, il vous suffit de prendre rendez-vous via notre outil RESERVIO.
                    Pour accéder à l’outil de réservation, merci de prendre connaissance du mail « AVIS EXPEDITION ».

                    Concernant l'installation de vos logiciels, merci de nous contacter au : 02 99 04 92 28.
                    Merci de revenir vers nous dans les 7 jours suivant la réception du PC, si ce délai est dépassé merci de faire un ticket EasyVista.
                    
                    Cordialement,
                    Postes de Travail France
                    """
            
            # Copier le texte dans le presse-papiers
            root.clipboard_clear()
            root.clipboard_append(text)
            root.update()

    messagebox.showinfo("Information", "Message copié dans le presse-papiers")


def reload_excel_file_btn():
    load_excel_file_endowment(hp_file_path)


# --- Getion excel Relance ---
def check_columns(df):
    # Nettoyer les noms de colonnes pour enlever les espaces inutiles
    df.columns = df.columns.str.strip()

    # Les noms de colonnes attendus
    required_columns = ['Mail', 'Matériel concerné', 'Etat rétro', 'Etape Relance', 'Mail', 'Mail', 'Mail Responsable']  # Vérifiez l'exactitude des noms

    # Vérifier si toutes les colonnes requises sont présentes
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        messagebox.showerror("Erreur", f"Colonnes manquantes : {', '.join(missing_columns)}")
        return False
    return True

def load_excel_file(file_path):
    if not file_path:
        messagebox.showerror("Erreur", "Aucun fichier sélectionné.")
        return
    file_path_retro = file_path 
    file_path_retro_cp = os.path.join(config_directory, 'tmp_dotation.xlsx')
    shutil.copy2(file_path_retro, file_path_retro_cp)
    path_excel_label_retro.config(text=file_path_retro.split("/")[-1])
    try: 
        
        df = pd.read_excel(file_path_retro_cp, sheet_name='Rétro')
        
        if df.empty:
            raise ValueError("Le fichier Excel est vide.")
        
        check_columns(df)
        update_email_listbox('excel')
        
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du chargement du fichier Excel : {e}")



# Fonction Maj Treeview Relance
def update_email_listbox(event):
    global df
    if 'df' not in globals() or df.empty:
        messagebox.showerror("Erreur", "Veuillez d'abord charger un fichier Excel valide.")
        return

    selected_preset = preset_combobox.get() 
    if selected_preset:
        try:
            if len(selected_preset.split()) < 2:
                messagebox.showerror("Erreur", "Le format du preset sélectionné est incorrect.")
                return

            # Récupère l'étape 
            etape = selected_preset.split()[1]  
            
            
            df['Etape Relance'] = df['Etape Relance'].str.strip()

            # Vérifie si les colonnes nécessaires sont présentes dans le fichier Excel
            if all(col in df.columns for col in ['Matériel concerné', 'Mail', 'Etat rétro', 'Etape Relance']):
                
                # Filtre les données en fonction de l'étape et des emails non terminés
                filter_conditions = (
                    (df['Etat rétro'] != 'Terminé') & 
                    (df['Etape Relance'] == f"Etape {etape}")
                )
                
                filtered_columns = df.loc[filter_conditions, ['Matériel concerné', 'Mail', 'Etat rétro', 'Etape Relance', 'Mail Responsable']]
                
                # Supprime les lignes avec des valeurs manquantes
                filtered_columns.dropna(subset=['Matériel concerné', 'Mail', 'Etape Relance'], inplace=True)
                
                # Efface les anciennes données du Treeview
                treeview_relance.delete(*treeview_relance.get_children())
                
                # Ajoute les nouvelles données filtrées dans le Treeview
                for _, row in filtered_columns.iterrows():
                    treeview_relance.insert("", "end", values=(
                        row['Matériel concerné'],
                        row['Mail'],
                        row['Mail Responsable'],
                        row['Etape Relance']
                    ))
            else:
                messagebox.showerror("Erreur", "Les colonnes nécessaires ne sont pas présentes dans le fichier Excel.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la mise à jour de la liste des emails : {e}")
    else:
        messagebox.showerror("Erreur", "Veuillez sélectionner un preset valide.")

def reload_excel_file_relance_btn():
    load_excel_file(hp_file_path)






# --- Envoie des emails relance ---
def send_emails():
    global df

    # Vérifie que le fichier Excel a bien été chargé
    if 'df' not in globals() or df.empty:
        messagebox.showerror("Erreur", "Veuillez d'abord charger un fichier Excel valide.")
        return

    # Récupérer les indices sélectionnés dans le Treeview
    selected_items = treeview_relance.selection()
    if not selected_items:
        messagebox.showerror("Erreur", "Aucun email sélectionné")
        return

    # Récupérer les emails sélectionnés depuis le Treeview
    selected_emails = [treeview_relance.item(item, 'values') for item in selected_items]

    # Configurer le serveur SMTP
    server, from_address = config_smtp()

    # Récupérer le preset de relance sélectionné
    preset_relance = presets.get(preset_combobox.get())
    if not preset_relance:
        messagebox.showerror("Erreur", "Preset de mail introuvable.")
        return

    subject, body_template_path = preset_relance[1], preset_relance[2]

    try:
        # Lecture du modèle d'email
        with open(body_template_path, 'r', encoding='utf-8') as file:
            body_template = file.read()
    except (UnicodeDecodeError, FileNotFoundError) as e:
        messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier de modèle d'email : {e}")
        server.quit()
        return

    # Remplace les accolades pour éviter les erreurs dans le template
    body_template = escape_curly_braces(body_template)

    for entry in selected_emails:
        # Extraction des informations depuis le Treeview
        materiel, mail_to, mail_responsable, etape = entry

        # Vérification des données et initialisation des variables si nécessaires
        if not mail_to or not isinstance(mail_to, str):
            messagebox.showerror("Erreur", "Adresse email invalide.")
            continue

        try:
            # Personnalisation du modèle avec les données de l'utilisateur
            body = body_template.format(
                materiel=materiel,
                etape=etape
            )
        except KeyError as e:
            messagebox.showerror("Erreur", f"Variable non définie dans le modèle d'email : {e}")
            continue  # Passer à l'entrée suivante

        # Création du message MIME
        mimemsg = MIMEMultipart()
        mimemsg['From'] = from_address
        mimemsg['To'] = mail_to.strip()
        if var1.get() == 1:  # Si la case responsable en CC est cochée
            mimemsg['Cc'] = mail_responsable.strip()
        mimemsg['Subject'] = subject
        mimemsg.attach(MIMEText(body, 'html', 'utf-8'))

        # Ajout de la pièce jointe (si présente)
        if attachment_file_path:
            try:
                with open(attachment_file_path, "rb") as attachment:
                    part = MIMEApplication(attachment.read(), _subtype="pdf")
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_file_path))
                    mimemsg.attach(part)
            except FileNotFoundError:
                messagebox.showerror("Erreur", "Pièce jointe introuvable")
                continue

        # Ajout des images inline dans l'email
        output_folder = os.path.dirname(body_template_path)
        image_files = [f for f in os.listdir(output_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]
        for image_file in image_files:
            with open(os.path.join(output_folder, image_file), 'rb') as img:
                mime_image = MIMEImage(img.read())
                mime_image.add_header('Content-ID', f'<{image_file}>')
                mime_image.add_header('Content-Disposition', 'inline', filename=image_file)
                mimemsg.attach(mime_image)

        # Envoi de l'email
        try:
            server.send_message(mimemsg)
            print(f"Email envoyé à {mail_to}")
            log_action(f"{formatted_datetime} Email envoyé à {mail_to} pour {materiel}")
            read_file_log()

        except smtplib.SMTPRecipientsRefused as e:
            print(f"Erreur d'envoi à {mail_to}: {e}")
            messagebox.showinfo("Erreur", f"Erreur d'envoi à {mail_to}: {e}")

    # Mise à jour du fichier Excel après envoi
    load_excel_file(hp_file_path)
    update_email_listbox('Maj')

    # Fermeture de la connexion au serveur SMTP
    server.quit()
    messagebox.showinfo("Succès", "Emails envoyés avec succès")

    


# --- Envoie des emails Doation/Remplacement ---
def send_emails_endowment():
    global old_dst
    # Récupérer les indices sélectionnés dans le Treeview
    selected_items = treeview.selection()
    if not selected_items:
        messagebox.showerror("Erreur", "Aucun email sélectionné")
        return

    selected_emails_endowment = [treeview.item(item, 'values') for item in selected_items]

    server, from_address = config_smtp()

    for entry in selected_emails_endowment:
        demandeur, mail_to, beneficiare, ville, expe, dst, old_dst = entry[:7]
        
        # Nouvelle condition pour choisir le modèle en fonction de "Ref exp"
        if expe.startswith("X"):
            preset_dotation = presets.get('dota/rempla chrono')  # Si "Ref exp" commence par X, choisir dota/rempla
        else:
            preset_dotation = presets.get('dota/rempla dpd')  # Sinon, utiliser un autre modèle

        subject, body_template_path = preset_dotation[1], preset_dotation[2]

        try:
            with open(body_template_path, 'r', encoding='utf-8') as file:
                body_template = file.read()
        except (UnicodeDecodeError, FileNotFoundError) as e:
            messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier de modèle d'email : {e}")
            server.quit()
            return

        body_template = escape_curly_braces(body_template)

    for entry in selected_emails_endowment:
        # Le Treeview renvoie une liste de valeurs correspondant aux colonnes
        demandeur, mail_to, beneficiare, ville, expe, dst, old_dst = entry

        if old_dst.strip().lower() == 'dotation':
            old_dst = '/'

        behavior = "some_default_value"  # Modifie ou ajuste cette valeur selon les besoins

        try:
            body = body_template.format(
                demandeur=demandeur,
                ville=ville,
                expe=expe,
                dst=dst,
                old_dst=old_dst,
                beneficiare=beneficiare,
                behavior=behavior
            )
        except KeyError as e:
            messagebox.showerror("Erreur", f"Variable non définie dans le modèle d'email : {e}")
            server.quit()
            return

        if not demandeur or demandeur == "nan":
            demandeur = mail_to

        if not demandeur.strip():
            continue

        mimemsg = MIMEMultipart()
        mimemsg['From'] = from_address
        mimemsg['To'] = demandeur.strip()
        mimemsg['Cc'] = mail_to.strip()
        mimemsg['Subject'] = subject
        mimemsg.attach(MIMEText(body, 'html', 'utf-8'))

        if attachment_file_path:
            with open(attachment_file_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), _subtype="pdf")
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_file_path))
                mimemsg.attach(part)

        output_folder = os.path.dirname(body_template_path)
        image_files = [f for f in os.listdir(output_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]
        for image_file in image_files:
            with open(os.path.join(output_folder, image_file), 'rb') as img:
                mime_image = MIMEImage(img.read())
                mime_image.add_header('Content-ID', f'<{image_file}>')
                mime_image.add_header('Content-Disposition', 'inline', filename=image_file)
                mimemsg.attach(mime_image)

        try:
            server.send_message(mimemsg)
            print(f"Email envoyé à {mail_to}")
            log_action(f"{formatted_datetime} Email envoyé à {mail_to} {ville}")
            read_file_log()

        except smtplib.SMTPRecipientsRefused as e:
            messagebox.showinfo("Erreur", f"Erreur d'envoi à {mail_to}: {e}")

    server.quit()
    messagebox.showinfo("Succès", "Emails envoyés avec succès")




# --- Gestion des presets ---
def show_html_preview():
    selected_preset = preset_listbox.get(tk.ACTIVE)
    if selected_preset:
        preset_name = " ".join(selected_preset.split(" ")[:-1])
        if preset_name in presets and presets[preset_name][0] == "HTML":
            webbrowser.open(presets[preset_name][2])
        else:
            messagebox.showerror("Erreur", "Veuillez sélectionner un preset HTML valide pour l'aperçu.")
    else:
        messagebox.showerror("Erreur", "Aucun preset sélectionné pour l'aperçu.")

# Fonction pour  les accolades HTML
def escape_curly_braces(content):
    content = content.replace("{", "{{").replace("}", "}}")
    content = content.replace("{{materiel}}", "{materiel}").replace("{{etape}}", "{etape}").replace("{{dst}}","{dst}").replace(
        "{{beneficiare}}", "{beneficiare}").replace("{{ville}}", "{ville}").replace("{{expe}}", "{expe}").replace(
        "{{old_dst}}", "{old_dst}")
    return content



def edit_preset_window(preset_name=None, is_html=False):
    def save_preset():
        subject = subject_entry_popup.get().strip()
        if is_html:
            content = mail_text.get()
        else:
            content = text.get("1.0", tk.END).strip()
        if preset_name:
            presets[preset_name] = ("HTML" if is_html else "Texte", subject, content)
        else:
            new_preset_name = name_entry.get().strip()
            if new_preset_name in presets:
                messagebox.showerror("Erreur", "Ce nom de preset existe déjà. Veuillez en choisir un autre.")
                return
            if new_preset_name:
                presets[new_preset_name] = ("HTML" if is_html else "Texte", subject, content)
        preset_combobox['values'] = list(presets.keys())
        preset_listbox.delete(0, tk.END)
        for preset in presets.keys():
            preset_listbox.insert(tk.END, f"{preset} ({presets[preset][0]})")
        write_config()
        window.destroy()

    def select_html_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            mail_text.delete(0, tk.END)
            mail_text.insert(0, file_path)

    window = Toplevel(root)
    window.title("Modifier le Preset" if preset_name else "Ajouter un Preset")
    window.geometry("400x400")

    if not preset_name:
        ttk.Label(window, text="Nom du Preset:").pack(pady=5)
        name_entry = ttk.Entry(window, width=50)
        name_entry.pack(pady=5)

    ttk.Label(window, text="Sujet de l'email:").pack(pady=5)
    subject_entry_popup = ttk.Entry(window, width=50)
    subject_entry_popup.pack(pady=5)

    if is_html:
        ttk.Label(window, text="Chemin du fichier HTML:").pack(pady=5)
        mail_text = ttk.Entry(window, width=50)
        mail_text.pack(pady=5)
        if preset_name:
            mail_text.insert(0, presets[preset_name][2])
            subject_entry_popup.insert(0, presets[preset_name][1])
        ttk.Button(window, text="Sélectionner un fichier", command=select_html_file).pack(pady=5)
    else:
        ttk.Label(window, text="Contenu du Preset:").pack(pady=5)
        text = Text(window, wrap=WORD, height=10, width=50)
        text.pack(pady=5)
        if preset_name:
            text.insert("1.0", presets[preset_name][2])
            subject_entry_popup.insert(0, presets[preset_name][1])

    ttk.Button(window, text="Sauvegarder", command=save_preset).pack(pady=10)


def add_preset(is_html=False):
    edit_preset_window(is_html=is_html)


def delete_preset():
    selected_preset = preset_listbox.get(tk.ACTIVE)
    preset_name = " ".join(selected_preset.split(" ")[:-1])
    if preset_name in presets:
        del presets[preset_name]
        preset_combobox['values'] = list(presets.keys())
        preset_listbox.delete(tk.ACTIVE)
        write_config()


def edit_preset():
    selected_preset = preset_listbox.get(tk.ACTIVE)
    preset_name = " ".join(selected_preset.split(" ")[:-1])
    if preset_name in presets:
        edit_preset_window(preset_name, is_html=presets[preset_name][0] == "HTML")


def load_attachment():
    global attachment_file_path
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        attachment_file_path = file_path
        attachment_label.config(text=os.path.basename(attachment_file_path))
        log_action(f"{formatted_datetime} Pièce jointe chargée : {file_path.split('/')[-1]}")
        read_file_log()
        write_config() 



def load_dotation_hp():
    global hp_file_path
    file_path = filedialog.askopenfilename()
    if file_path:
        hp_file_path = file_path
        hp_file_label.config(text=os.path.basename(hp_file_path))
        log_action(f"{formatted_datetime} Fichier HP chargée : {file_path.split('/')[-1]}")
        read_file_log()
        write_config() 




# --- Gestion interface graphique ---
root = tk.Tk()

var_dpd = tk.IntVar()
var_chrono = tk.IntVar()

style = Style(theme='superhero')

root.title("Mail Sender")
root.geometry("1200x1200")
root.resizable(False, False)

# Lire la version
version = read_version()

# label pour afficher la version
version_label = ttk.Label(root, text=f"Version {version}", anchor='e')
version_label.pack(side=tk.BOTTOM, anchor='e', padx=10, pady=10)

root.iconbitmap('Mail_Sender\\logo.ico')

notebook = ttk.Notebook(root)
notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

#### Onglet Informations de Connexion ####
frame_connexion = ttk.Frame(notebook, padding=20)
notebook.add(frame_connexion, text="Informations de Connexion")

ttk.Label(frame_connexion, text="Nom d'utilisateur:").pack(pady=5, anchor=W)
username_entry = ttk.Entry(frame_connexion, width=50)
username_entry.pack(pady=5)

ttk.Label(frame_connexion, text="Mot de passe:").pack(pady=5, anchor=W)
password_entry = ttk.Entry(frame_connexion, width=50, show="*")
password_entry.pack(pady=5)

ttk.Label(frame_connexion, text="Adresse d'envoi:").pack(pady=5, anchor=W)
from_address_entry = ttk.Entry(frame_connexion, width=50)
from_address_entry.pack(pady=5)

ttk.Button(frame_connexion, text="Enregistrer", command=write_config, style='success.TButton').pack(pady=10)

#### Onglet Relance ####
frame_email = ttk.Frame(notebook, padding=20)
notebook.add(frame_email, text="Relance")

button_frame = ttk.Frame(frame_email)
button_frame.pack(pady=10)

ttk.Button(button_frame, text="Charger fichier Excel", command=reload_excel_file_relance_btn, style='success.TButton').pack(side=LEFT,padx=10,pady=10)

path_excel_label_retro = ttk.Label(frame_email, text="Aucun fichier chargé")
path_excel_label_retro.pack()

# Filtrer les clés  pour mail 1, mail 2, mail 3
filtered_presets = [key for key in ["mail 1", "mail 2", "mail 3"] if key in presets]


#combobox pour l'onglet Relance
ttk.Label(frame_email, text="Sélectionnez un preset de mail:").pack(pady=5, anchor=W)
preset_combobox = ttk.Combobox(frame_email, values=filtered_presets, state="readonly")
preset_combobox.pack(pady=5)


var1 = tk.IntVar()
cc_responsable = ttk.Checkbutton(frame_email, text='Responsable en Cc', variable=var1, onvalue=1, offvalue=0)
cc_responsable.pack()

style = ttk.Style()
style.configure("Treeview.Heading", foreground="white", background="green", font=("Helvetica", 10, "bold"))


treeview_relance = ttk.Treeview(frame_email, columns=("DST", "Mail User", "Mail Responsable", "Etape"), show="headings", selectmode="extended")
treeview_relance.pack(fill="both", expand=True)

# Configurer les en-têtes des colonnes avec alignement à gauche et style
treeview_relance.heading("DST", text="DST", anchor="w")
treeview_relance.heading("Mail User", text="Mail User", anchor="w")
treeview_relance.heading("Mail Responsable", text="Mail Responsable", anchor="w")
treeview_relance.heading("Etape", text="Etape", anchor="w")

# Configurer la largeur des colonnes
treeview_relance.column("DST", width=170, anchor="w")
treeview_relance.column("Mail User", width=170, anchor="w")
treeview_relance.column("Mail Responsable", width=150, anchor="w")
treeview_relance.column("Etape", width=100, anchor="w")

# Liaison de la sélection du combobox
preset_combobox.bind("<<ComboboxSelected>>", update_email_listbox)


# Màj dynamiquement les valeurs après l'exécution de read_config()
def update_combobox_values():
    filtered_presets = [key for key in ["mail 1", "mail 2", "mail 3"] if key in presets]
    preset_combobox['values'] = filtered_presets



button_frame_email = ttk.Frame(frame_email)
button_frame_email.pack(pady=10)

ttk.Button(button_frame_email, text="Envoyer les emails", command=send_emails, style='primary.TButton').pack(side=LEFT, padx=10)

##### Onglet Dotation/Rempla #####
frame_endowment = ttk.Frame(notebook, padding=20)
notebook.add(frame_endowment, text="Dotation/Rempla")

button_frame_endowment = ttk.Frame(frame_endowment)
button_frame_endowment.pack(pady=10)

path_excel_label = ttk.Label(frame_endowment, text="Aucun fichier chargé")
path_excel_label.pack()

ttk.Button(button_frame_endowment, text="Charger fichier Excel", command=import_excel_file_endowment,style='success.TButton').pack(side=LEFT, padx=10, pady=5)

#img bouton reload
img_reload = tk.PhotoImage(file=r"Mail_Sender\src\img\reload.png")
photoimage = img_reload.subsample(20, 20) # modifier la taille
ttk.Button(button_frame_endowment, image=photoimage, command=reload_excel_file_btn, style="success.Outline.TButton").pack()


style = ttk.Style()
style.configure("Treeview.Heading", foreground="white", background="green", font=("Helvetica", 10, "bold"))


treeview = ttk.Treeview(frame_endowment, columns=("demandeur", "mail", "nom_prenom", "ville", "ref_exp", "dst", "ancien_dst"), show="headings", selectmode="extended")
treeview.pack(fill="both", expand=True)

# Configurer les en-têtes des colonnes avec alignement à gauche et style
treeview.heading("demandeur", text="Demandeur", anchor="w")
treeview.heading("mail", text="Mail", anchor="w")
treeview.heading("nom_prenom", text="Nom, Prénom", anchor="w")
treeview.heading("ville", text="Ville", anchor="w")
treeview.heading("ref_exp", text="Réf exp", anchor="w")
treeview.heading("dst", text="DST", anchor="w")
treeview.heading("ancien_dst", text="Ancien DST", anchor="w")

# Configurer la largeur des colonnes
treeview.column("demandeur", width=170, anchor="w")
treeview.column("mail", width=170, anchor="w")
treeview.column("nom_prenom", width=150, anchor="w")
treeview.column("ville", width=100, anchor="w")
treeview.column("ref_exp", width=120, anchor="w")
treeview.column("dst", width=100, anchor="w")
treeview.column("ancien_dst", width=50, anchor="w")

boutton_frame_type = ttk.Frame(frame_endowment)
boutton_frame_type.pack(pady=10)

button_frame_email_1 = ttk.Frame(frame_endowment)
button_frame_email_1.pack(pady=10)


ttk.Button(button_frame_email_1, text="Envoyer les emails", command=send_emails_endowment,style='primary.TButton').pack(side=LEFT, padx=10)
ttk.Button(button_frame_email_1, text="Copier message", command=copy_selected_email, style='primary.TButton').pack(side=LEFT, padx=10)



# Onglet Liste des presets
frame_msg = ttk.Frame(notebook, padding=20)
notebook.add(frame_msg, text="Presets Mails")

mail_labels = {}

preset_listbox = tk.Listbox(frame_msg, height=20)
preset_listbox.pack(pady=5, fill=tk.X)

preset_button_frame = ttk.Frame(frame_msg)
preset_button_frame.pack(pady=10)

ttk.Button(preset_button_frame, text="+ (Texte)", command=lambda: add_preset(is_html=False),style='success.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="+ (HTML)", command=lambda: add_preset(is_html=True),style='success.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="-", command=delete_preset, style='danger.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="Modifier", command=edit_preset, style='warning.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="Aperçu HTML", command=show_html_preview, style='info.TButton').pack(side=LEFT,padx=5)

ttk.Button(frame_msg, text="Charger Pièce Jointe (PDF)", command=load_attachment, style='success.TButton').pack(pady=10)
attachment_label = ttk.Label(frame_msg, text="Aucune pièce jointe chargée")
attachment_label.pack(pady=5)

ttk.Button(frame_msg, text="Charger fichier HP", command=load_dotation_hp, style='success.TButton').pack(pady=10)
hp_file_label = ttk.Label(frame_msg, text="Aucune pièce jointe chargée")
hp_file_label.pack(pady=5)

#Onglet action récentes
frame_action = ttk.Frame(notebook, padding=20)
notebook.add(frame_action, text="Actions Recentes")

button_frame_log = ttk.Frame(frame_action)
button_frame_log.pack(pady=10)


#btn delete log
ttk.Button(button_frame_log, text="Vider logs", command=delete_log_file, style="sucess.TButton").pack()


def read_file_log():
    with open(log_file, 'a') as f:
        pass
    if os.path.exists(log_file):
        with open(log_file, "r", encoding="utf-8") as file:
            content = file.read()
            read_log.delete("1.0", tk.END)
            read_log.insert("1.0", content)



read_log = tk.Text(frame_action, height=50, width=70)
read_log.pack()
read_file_log()

read_config()
update_combobox_values()

root.mainloop()
