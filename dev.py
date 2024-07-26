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
import extract_msg
import re

dota_file = None
filtered_df = None


def sanitize_path(path):
    # Remove extra spaces and normalize the path
    return os.path.normpath(path.strip())


def convert_msg_to_html(msgFolder, msgOutputFolder, msgFile):
    try:
        # Charger le fichier .msg
        msg = extract_msg.Message(os.path.join(msgFolder, msgFile))
        msgName = os.path.splitext(msgFile)[0]

        output_html_file = f"{msgName}.html"

        # Récupérer le corps du message au format HTML
        htmlBody = msg.htmlBody
        if not htmlBody:
            raise ValueError("Le fichier .msg ne contient pas de corps HTML.")

        # Décoder le contenu HTML si nécessaire
        if isinstance(htmlBody, bytes):
            htmlBody = htmlBody.decode('utf-8', errors='ignore')

        # Créer le dossier de sortie s'il n'existe pas
        outputFolder = sanitize_path(os.path.join(msgOutputFolder, msgName))
        if not os.path.exists(outputFolder):
            os.makedirs(outputFolder)

        # Remplacer les CIDs par les noms de fichiers des pièces jointes
        searchPattern = re.compile(r'src="cid:([^"]+)"')
        attachments_dict = {re.sub(r'[@<>]', '', attachment.cid): attachment for attachment in msg.attachments if
                            attachment.cid}

        # Trier les CIDs trouvés dans le HTML
        cids_in_html = searchPattern.findall(htmlBody)
        for cid in cids_in_html:
            cleaned_cid = re.sub(r'[@<>]', '', cid)
            if cleaned_cid in attachments_dict:
                attachment = attachments_dict[cleaned_cid]
                imageName = attachment.longFilename or attachment.shortFilename
                attachment_path = sanitize_path(os.path.join(outputFolder, imageName))
                with open(attachment_path, 'wb') as f:
                    f.write(attachment.data)

                # Remplacer les CIDs dans le HTML
                htmlBody = htmlBody.replace(f'src="cid:{cid}"', f'src="{imageName}"')

        # Sauvegarder le HTML dans un fichier
        output_html_path = sanitize_path(os.path.join(outputFolder, output_html_file))
        with open(output_html_path, 'w', encoding='utf-8') as htmlFile:
            htmlFile.write(htmlBody)

        print(f"{output_html_file} : done")

    except Exception as e:
        print(f"Error: {str(e)}")


def select_msg_file():
    msgFilePath = filedialog.askopenfilename(title="Sélectionnez un fichier .msg", filetypes=[("MSG Files", "*.msg")])
    if msgFilePath:
        msgFolder = os.path.dirname(msgFilePath)
        msgOutputFolder = os.path.join(msgFolder, "output")
        if not os.path.exists(msgOutputFolder):
            os.makedirs(msgOutputFolder)
        convert_msg_to_html(msgFolder, msgOutputFolder, os.path.basename(msgFilePath))


# Liste des presets de mail
presets = {}

# Jour et heure
current_datetime = datetime.now()
formatted_datetime = current_datetime.strftime("%d/%m/%Y %H:%M:%S")

# Fichier de configuration
config_directory = os.path.join(os.path.expanduser("~"), ".retro_app")
os.makedirs(config_directory, exist_ok=True)
config_file = os.path.join(config_directory, 'config.ini')
log_file = os.path.join(config_directory, 'log.txt')
loaded_file_path = None  # Variable pour stocker le chemin du fichier chargé
attachment_file_path = None  # Variable pour stocker le chemin de la pièce jointe


# Lire la configuration
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

    preset_combobox['values'] = list(presets.keys())
    preset_combobox_endowment['values'] = list(presets.keys())
    if 'dotation' not in presets:
        messagebox.showerror("Erreur", "Le preset 'Dotation' n'existe pas. Veuillez le créer avant de continuer.")


# Écrire la configuration
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

    with open(config_file, 'w') as configfile:
        config.write(configfile)
    print("Configuration écrite:", presets)


def log_action(action):
    with open(log_file, "a", encoding="utf-8") as log_file_obj:
        log_file_obj.write(action + "\n")


# Dotation
def load_excel_file_endowment():
    print(var1_2.get())
    print(var1_1.get())
    if var1_1.get() == 0 and var1_2.get() == 0:
        messagebox.showerror("Erreur", "Aucun type sélectionné")
        return

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    global dota_file
    file_path_endowment = file_path  # Enregistrer le chemin du fichier chargé
    dota_file = os.path.join(config_directory, 'tmp_dotation.xlsx')
    print(dota_file)
    shutil.copy2(file_path_endowment, dota_file)
    path_excel_label.config(text=file_path_endowment.split("/")[-1])
    read_excel_file_endowment(dota_file)


def read_excel_file_endowment(file_path):
    global filtered_df  # Assurez-vous que filtered_df est global
    df = pd.read_excel(file_path, sheet_name='Agence')
    df.columns = df.iloc[0]
    df = df[1:]
    selected_columns = df[[
        'Demandeur', 'NOM,Prénom', 'Mail', 'Ancien DST', 'Ville', 'DST', 'Ref exp', 'Rempla/ Dotation', 'Logiciels']]
    email_listbox_endowment.delete(0, tk.END)  # Utiliser le widget Listbox pour afficher les emails
    if var1_1.get() == 1:
        filtered_df = selected_columns[(
            selected_columns['Rempla/ Dotation'].str.contains('Dotation', case=False, na=False)) & (
            selected_columns['Logiciels'].str.contains('Mail', case=False, na=False))]
    else:
        filtered_df = selected_columns[(
            selected_columns['Rempla/ Dotation'].str.contains('Rempla', case=False, na=False)) & (
            selected_columns['Logiciels'].str.contains('Mail', case=False, na=False))]

    for index, row in filtered_df.iterrows():
        email_text = f"{row['Demandeur']} - {row['Mail']} - {row['NOM,Prénom']} - {row['Ville']} - {row['Ref exp']} - {row['DST']}"
        if var1_2.get() == 1:
            email_text += f" - {row['Ancien DST']}"

        # Ajouter le texte dans le widget Listbox
        email_listbox_endowment.insert(tk.END, email_text)

def copy_selected_email():
    selected_indices = email_listbox_endowment.curselection()
    if not selected_indices:
        messagebox.showerror("Erreur", "Aucun email sélectionné")
        return

    for i in selected_indices:
        email_text = email_listbox_endowment.get(i)
        parts = email_text.split(" - ")
        if len(parts) >= 6:
            demandeur, mail_to, beneficiare, ville, expe, dst = parts[:6]
            old_dst = parts[6] if len(parts) == 7 else ""

            text = f"""
                    Bonjour,
                    
                    Votre PC a été expédié. ({dst})
                    Lieu : {ville}
                    
                    Cordialement,
                    Postes de Travail France
                    """
            root.clipboard_clear()
            root.clipboard_append(text)
            root.update()  # now it stays on the clipboard after the window is closed
            messagebox.showinfo("Information", "Texte copié dans le presse-papiers")


# Envoyer emails retro
def send_emails():
    selected_indices = email_listbox.curselection()
    if not selected_indices:
        messagebox.showerror("Erreur", "Aucun email sélectionné")
        return

    selected_emails = [email_listbox.get(i) for i in selected_indices]

    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    username = username_entry.get()
    password = password_entry.get()
    from_address = from_address_entry.get()

    write_config()

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)

    subject = subject_entry.get()
    body_template = body_text.get("1.0", tk.END)

    for i, entry in enumerate(selected_emails):
        materiel, mail_to, mail_responsable, etape = entry.split(" - ")

        try:
            body = body_template.format(materiel=materiel, etape=etape, mail_responsable=mail_responsable)
        except KeyError as e:
            messagebox.showerror("Erreur", f"Variable non définie dans le modèle d'email : {e}")
            server.quit()
            return

        mimemsg = MIMEMultipart()
        mimemsg['From'] = from_address
        mimemsg['To'] = mail_to.strip()

        if var1.get() == 1:
            mimemsg['Cc'] = mail_responsable.strip()

        mimemsg['Subject'] = subject
        mimemsg.attach(MIMEText(body, 'html'))

        if attachment_file_path:
            with open(attachment_file_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), _subtype="pdf")
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_file_path))
                mimemsg.attach(part)

        try:
            server.send_message(mimemsg)
            print(f"Email envoyé à {mail_to}")
            log_action(f"{formatted_datetime} Email envoyé à {mail_to} {materiel}")
            read_file_log()

            df.loc[(df['Matériel concerné'] == materiel) & (df['Mail'] == mail_to), 'Etape'] = f"Etape {int(etape.split()[1]) + 1}"

        except smtplib.SMTPRecipientsRefused as e:
            print(f"Erreur d'envoi à {mail_to}: {e}")

    server.quit()
    df.to_excel(loaded_file_path, index=False)
    messagebox.showinfo("Succès", "Emails envoyés avec succès et étapes mises à jour")
    update_email_listbox()


# Envoyer les emails dotation
def send_emails_endowment():
    global old_dst
    selected_indices_endowment = email_listbox_endowment.curselection()
    if not selected_indices_endowment:
        messagebox.showerror("Erreur", "Aucun email sélectionné")
        return

    selected_emails_endowment = [email_listbox_endowment.get(i) for i in selected_indices_endowment]

    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    username = username_entry.get()
    password = password_entry.get()
    from_address = from_address_entry.get()

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)

    if var1_1.get() == 1:
        preset_dotation = presets.get('dotation')
    else:
        preset_dotation = presets.get('rempla')

    if not preset_dotation:
        messagebox.showerror("Erreur", "Le preset 'Dotation' n'existe pas. Veuillez le créer avant de continuer.")
        server.quit()
        return

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
        if var1_1.get() == 1:
            demandeur, mail_to, beneficiare, ville, expe, dst = entry.split(" - ")
            print(var1_1.get())
        else:
            demandeur, mail_to, beneficiare, ville, expe, dst, old_dst = entry.split(" - ")

        behavior = "some_default_value"

        try:
            if var1_1.get() == 1:
                body = body_template.format(
                    demandeur=demandeur,
                    ville=ville,
                    expe=expe,
                    dst=dst,
                    beneficiare=beneficiare,
                    behavior=behavior
                )
            else:
                body = body_template.format(
                    demandeur=demandeur,
                    ville=ville,
                    expe=expe,
                    dst=dst,
                    beneficiare=beneficiare,
                    old_dst=old_dst,
                    behavior=behavior
                )
        except KeyError as e:
            messagebox.showerror("Erreur", f"Variable non définie dans le modèle d'email : {e}")
            server.quit()
            return

        if not demandeur or demandeur == "nan":
            demandeur = mail_to

        if not demandeur.strip():
            continue  # Skip sending email if demandeur is empty after fallback

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
            print(f"Erreur d'envoi à {mail_to}: {e}")

    server.quit()
    messagebox.showinfo("Succès", "Emails envoyés avec succès")


# Fichier Excel retro
def load_excel_file():
    global df, loaded_file_path
    file_path = filedialog.askopenfilename()
    if not file_path:
        return
    df = pd.read_excel(file_path)
    loaded_file_path = file_path
    path_excel_label_retro.config(text=loaded_file_path.split("/")[-1])
    update_email_listbox()


def update_email_listbox():
    global df
    if 'df' not in globals():
        messagebox.showerror("Erreur", "Veuillez d'abord charger un fichier Excel.")
        return

    selected_preset = preset_combobox.get()
    if selected_preset:
        etape = selected_preset.split()[1]

        selected_columns = df[['Matériel concerné', 'Mail', 'Etat rétro', 'Etape', 'Mail Responsable']]
        filtered_columns = selected_columns[
            (selected_columns['Etat rétro'] != 'Terminé') & (selected_columns['Etape'] == f"Etape {etape}")]
        filtered_columns = filtered_columns.dropna(subset=['Matériel concerné', 'Mail', 'Etape'])

        email_listbox.delete(0, tk.END)
        for index, row in filtered_columns.iterrows():
            email_listbox.insert(tk.END,
                                 f"{row['Matériel concerné']} - {row['Mail']} - {row['Mail Responsable']} - {row['Etape']}")

# Fichier Mail (HTML)
def load_mail_file(mail_label, preset_name):
    file_path_mail = filedialog.askopenfilename(filetypes=[("HTML files", "*.htm;*.html")])
    if file_path_mail:
        mail_label.config(text=file_path_mail)
        presets[preset_name] = ("HTML", "", file_path_mail)
        write_config()

# Fonction pour afficher un aperçu du contenu HTML
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

# Fonction pour échapper les accolades dans le contenu HTML
def escape_curly_braces(content):
    content = content.replace("{", "{{").replace("}", "}}")
    content = content.replace("{{materiel}}", "{materiel}").replace("{{etape}}", "{etape}").replace("{{dst}}",
                                                                                                    "{dst}").replace(
        "{{beneficiare}}", "{beneficiare}").replace("{{ville}}", "{ville}").replace("{{expe}}", "{expe}").replace(
        "{{old_dst}}", "{old_dst}")
    return content

# Fonction pour mettre à jour le contenu de l'email en fonction du preset sélectionné
def update_body_text(event):
    selected_preset = preset_combobox.get()
    if selected_preset is not None and selected_preset in presets:
        subject_entry.delete(0, tk.END)
        subject_entry.insert(0, presets[selected_preset][1])

        body_text.delete("1.0", tk.END)
        content = presets[selected_preset][2]
        if presets[selected_preset][0] == "HTML":
            file_path = content
            if not file_path:
                messagebox.showerror("Erreur",
                                     f"Le chemin pour {selected_preset} est vide. Veuillez charger un fichier HTML.")
                return
            content = None
            for encoding in ['utf-8', 'latin-1', 'cp1252']:
                try:
                    with open(file_path, 'r', encoding=encoding) as file:
                        content = file.read()
                        break
                except UnicodeDecodeError:
                    continue
                except FileNotFoundError:
                    messagebox.showerror("Erreur", f"Le fichier {file_path} n'a pas été trouvé.")
                    return
            if content:
                content = escape_curly_braces(content)
                body_text.insert("1.0", content)
            else:
                messagebox.showerror("Erreur", "Impossible de lire le fichier avec les encodages disponibles")
        else:
            body_text.insert("1.0", content)

        if selected_preset == "mail 3":
            var1.set(1)
        else:
            var1.set(0)

        update_email_listbox()

# Fonction pour ajouter ou éditer un preset avec une fenêtre de texte
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
        file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.htm;*.html")])
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

# Ajouter nouveau preset
def add_preset(is_html=False):
    edit_preset_window(is_html=is_html)

# Supprimer preset
def delete_preset():
    selected_preset = preset_listbox.get(tk.ACTIVE)
    preset_name = " ".join(selected_preset.split(" ")[:-1])
    if preset_name in presets:
        del presets[preset_name]
        preset_combobox['values'] = list(presets.keys())
        preset_listbox.delete(tk.ACTIVE)
        write_config()

# Modifier preset
def edit_preset():
    selected_preset = preset_listbox.get(tk.ACTIVE)
    preset_name = " ".join(selected_preset.split(" ")[:-1])
    if preset_name in presets:
        edit_preset_window(preset_name, is_html=presets[preset_name][0] == "HTML")

# Fonction pour créer un widget pour charger un mail HTML
def create_mail_widget(parent, text):
    ttk.Label(parent, text=text).pack(pady=5, anchor=W)
    button_frame = ttk.Frame(parent)
    label = ttk.Label(parent, text="", wraplength=400)
    mail_labels[text] = label
    label.pack(pady=5, anchor=W)
    ttk.Button(button_frame, text="Charger Mail (HTML)", command=lambda: load_mail_file(label, text),
               style='success.TButton').pack(side=LEFT, pady=10)
    button_frame.pack()

def load_attachment():
    global attachment_file_path
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        attachment_file_path = file_path
        attachment_label.config(text=os.path.basename(attachment_file_path))
        log_action(f"{formatted_datetime} Pièce jointe chargée : {file_path.split('/')[-1]}")
        write_config()  # Enregistrer le chemin de la pièce jointe dans la configuration

# Interface graphique
root = tk.Tk()

var1_1 = tk.IntVar()
var1_2 = tk.IntVar()

style = Style(theme='superhero')

root.title("Mail Sender")
root.geometry("800x900")
root.resizable(False, False)

root.iconbitmap('logo.ico')

notebook = ttk.Notebook(root)
notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

# Onglet Informations de Connexion
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

# Onglet Relance
frame_email = ttk.Frame(notebook, padding=20)
notebook.add(frame_email, text="Relance")

button_frame = ttk.Frame(frame_email)
button_frame.pack(pady=10)

ttk.Button(button_frame, text="Charger fichier Excel", command=load_excel_file, style='success.TButton').pack(side=LEFT,
                                                                                                              padx=10,
                                                                                                              pady=10)

path_excel_label_retro = ttk.Label(frame_email, text="Aucun fichier chargé")
path_excel_label_retro.pack()

ttk.Label(frame_email, text="Sélectionnez un preset de mail:").pack(pady=5, anchor=W)
preset_combobox = ttk.Combobox(frame_email, values=list(presets.keys()), state="readonly")
preset_combobox.pack(pady=5)
preset_combobox.bind("<<ComboboxSelected>>", update_body_text)

var1 = tk.IntVar()
cc_responsable = ttk.Checkbutton(frame_email, text='Responsable en Cc', variable=var1, onvalue=1, offvalue=0)
cc_responsable.pack()

ttk.Label(frame_email, text="Sujet de l'email:").pack(pady=5, anchor=W)
subject_entry = ttk.Entry(frame_email, width=50)
subject_entry.pack(pady=5)

ttk.Label(frame_email, text="Contenu de l'email:").pack(pady=5, anchor=W)
body_text = tk.Text(frame_email, wrap=WORD, height=10, width=85)
body_text.pack(pady=5)

ttk.Label(frame_email, text="Sélectionnez les emails:").pack(pady=5, anchor=W)
ttk.Label(frame_email, text=" DST |  Mail User  |  Mail Responsable  |  Etape").pack(padx=55, anchor=W)
email_listbox = tk.Listbox(frame_email, selectmode=tk.MULTIPLE, height=10, width=85)
email_listbox.pack(pady=5)

button_frame_email = ttk.Frame(frame_email)
button_frame_email.pack(pady=10)

ttk.Button(button_frame_email, text="Envoyer les emails", command=send_emails, style='primary.TButton').pack(side=LEFT,
                                                                                                             padx=10)


# Onglet dotation/rempla


def on_dotation_change():
    # Si Dotation est cochée, désactiver Remplacement
    if var1_1.get() == 1:
        var1_2.set(0)  # Décocher Remplacement
        preset_combobox_endowment.set("dotation")
        read_excel_file_endowment(dota_file)
    else:
        var1_2.set(1)
        preset_combobox_endowment.set("rempla")


def on_rempla_change():
    # Si Remplacement est cochée, désactiver Dotation
    if var1_2.get() == 1:
        var1_1.set(0)  # Décocher Dotation
        preset_combobox_endowment.set("rempla")
        read_excel_file_endowment(dota_file)
    else:
        var1_1.set(1)
        preset_combobox_endowment.set("dotation")


frame_endowment = ttk.Frame(notebook, padding=20)
notebook.add(frame_endowment, text="Dotation/Rempla")

check_dota = ttk.Checkbutton(frame_endowment, text='Dotation', variable=var1_1, command=on_dotation_change, onvalue=1,
                             offvalue=0)
check_dota.pack(anchor=W, pady=10)

check_rempla = ttk.Checkbutton(frame_endowment, text='Remplacement', variable=var1_2, command=on_rempla_change,
                               onvalue=1, offvalue=0)
check_rempla.pack(anchor=W)

button_frame_endowment = ttk.Frame(frame_endowment)
button_frame_endowment.pack(pady=10)

path_excel_label = ttk.Label(frame_endowment, text="Aucun fichier chargé")
path_excel_label.pack()

ttk.Button(button_frame_endowment, text="Charger fichier Excel", command=load_excel_file_endowment,
           style='success.TButton').pack(side=LEFT, padx=10, pady=5)

ttk.Label(frame_endowment, text="Sélectionnez un preset de mail:").pack(pady=15, anchor=W)
preset_combobox_endowment = ttk.Combobox(frame_endowment, values=list(presets.keys()), state="readonly")
preset_combobox_endowment.pack(pady=5)
preset_combobox_endowment.config(state="disabled")

ttk.Label(frame_endowment, text="Sélectionnez les emails:").pack(pady=5, anchor=W)
email_listbox_endowment = tk.Listbox(frame_endowment, selectmode=tk.MULTIPLE, height=20, width=150)
email_listbox_endowment.pack(pady=5)

button_frame_email_1 = ttk.Frame(frame_endowment)
button_frame_email_1.pack(pady=10)

ttk.Button(button_frame_email_1, text="Envoyer les emails", command=send_emails_endowment,
           style='primary.TButton').pack(side=LEFT, padx=10)

ttk.Button(button_frame_email_1, text="Copier l'email sélectionné", command=copy_selected_email, style='primary.TButton').pack(side=LEFT, padx=10)

# Onglet Liste des presets
frame_msg = ttk.Frame(notebook, padding=20)
notebook.add(frame_msg, text="Presets Mails")

mail_labels = {}

preset_listbox = tk.Listbox(frame_msg, height=20)
preset_listbox.pack(pady=5, fill=tk.X)

preset_button_frame = ttk.Frame(frame_msg)
preset_button_frame.pack(pady=10)

ttk.Button(preset_button_frame, text="+ (Texte)", command=lambda: add_preset(is_html=False),
           style='success.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="+ (HTML)", command=lambda: add_preset(is_html=True),
           style='success.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="-", command=delete_preset, style='danger.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="Modifier", command=edit_preset, style='warning.TButton').pack(side=LEFT, padx=5)
ttk.Button(preset_button_frame, text="Aperçu HTML", command=show_html_preview, style='info.TButton').pack(side=LEFT,
                                                                                                          padx=5)

ttk.Button(frame_msg, text="Charger Pièce Jointe (PDF)", command=load_attachment, style='success.TButton').pack(pady=10)
attachment_label = ttk.Label(frame_msg, text="Aucune pièce jointe chargée")
attachment_label.pack(pady=5)

# Onglet action récentes
frame_action = ttk.Frame(notebook, padding=20)
notebook.add(frame_action, text="Actions Recentes")

#Onglet msg to html
frame_msg2html = ttk.Frame(notebook, padding=20)
notebook.add(frame_msg2html, text="MSG to HTML")

ttk.Button(frame_msg2html, text="Charger fichier MSG", command=select_msg_file, style='success.TButton').pack(pady=5)


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
print("OK")

read_config()

root.mainloop()
