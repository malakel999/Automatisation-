import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import schedule
import time

# Chemins des fichiers Excel
file1_path = 'Factures PR MFS 2024 4 check du 24 07 2024 (1).xlsx'
file2_path = 'Ventes et avoirs mi juillet 2024.xlsx'

# Charger les données à partir des feuilles des fichiers Excel
file1_encours_2024_df = pd.read_excel(file1_path, sheet_name='encours 2024')
file1_en_attente_df = pd.read_excel(file1_path, sheet_name='en attente')
file2_factures_ventes_df = pd.read_excel(file2_path, sheet_name='Factures ventes ')
file2_avoirs_ventes_df = pd.read_excel(file2_path, sheet_name='Avoirs ventes')

# Calculer les résumés des factures
total_factures_encours = file1_encours_2024_df['NO_FACT_FOUR'].nunique()
montant_total_encours = file1_encours_2024_df['MT_LIG_FIN'].sum()

total_factures_en_attente = file1_en_attente_df['N° Facture'].nunique()
montant_total_en_attente = file1_en_attente_df['Montant'].sum()

total_factures_ventes = file2_factures_ventes_df['N facture'].nunique()
montant_total_ventes = file2_factures_ventes_df['Montant'].sum()

total_avoirs_ventes = file2_avoirs_ventes_df['N document'].nunique()
montant_total_avoirs = file2_avoirs_ventes_df['montant'].sum()

# Définir la variable 'resume_factures'
resume_factures = {
    "Factures en cours": {
        "Nombre total": total_factures_encours,
        "Montant total": montant_total_encours
    },
    "Factures en attente": {
        "Nombre total": total_factures_en_attente,
        "Montant total": montant_total_en_attente
    },
    "Ventes facturées": {
        "Nombre total": total_factures_ventes,
        "Montant total": montant_total_ventes
    },
    "Avoirs des ventes": {
        "Nombre total": total_avoirs_ventes,
        "Montant total": montant_total_avoirs
    }
}

# Fonction pour envoyer l'email
def envoyer_rapport_hebdomadaire():
    # Configuration de l'email
    email_from = 'elbakalimalak312@gmail.com'
    email_to = 'elbakali.malak@etu.uae.ac.ma'
    email_subject = 'Résumé Hebdomadaire des Factures et Ventes'

    # Contenu du rapport
    rapport_texte = f"""
    Résumé Hebdomadaire des Factures :
    -----------------------
    Factures en cours :
    - Nombre total : {resume_factures['Factures en cours']['Nombre total']}
    - Montant total : {resume_factures['Factures en cours']['Montant total']:.2f}

    Factures en attente :
    - Nombre total : {resume_factures['Factures en attente']['Nombre total']}
    - Montant total : {resume_factures['Factures en attente']['Montant total']:.2f}

    Ventes facturées :
    - Nombre total : {resume_factures['Ventes facturées']['Nombre total']}
    - Montant total : {resume_factures['Ventes facturées']['Montant total']:.2f}

    Avoirs des ventes :
    - Nombre total : {resume_factures['Avoirs des ventes']['Nombre total']}
    - Montant total : {resume_factures['Avoirs des ventes']['Montant total']:.2f}
    """

    # Configuration du message
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = email_subject
    msg.attach(MIMEText(rapport_texte, 'plain'))

    # Configuration du serveur SMTP (exemple pour Gmail)
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'Malak El Bakali '
    smtp_password = '******'

    # Envoyer l'email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(email_from, email_to, msg.as_string())
        server.quit()
        print('Email hebdomadaire envoyé avec succès')
    except Exception as e:
        print(f'Erreur lors de l\'envoi de l\'email: {e}')

# Planifier l'envoi toutes les 10 secondes pour le test
schedule.every(10).seconds.do(envoyer_rapport_hebdomadaire)

# Garder le script en cours d'exécution
while True:
    schedule.run_pending()
    time.sleep(1)
