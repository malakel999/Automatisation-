import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import schedule
import time

# Chemins des fichiers Excel
file1_path = 'Factures PR MFS 2024 4 check du 24 07 2024 (1).xlsx'

# Charger les données à partir des feuilles des fichiers Excel
file1_en_attente_df = pd.read_excel(file1_path, sheet_name='en attente')

# Fonction pour envoyer une alerte si une facture est en attente trop longtemps
def verifier_factures_en_attente():
    # Définir la date de 3 mois auparavant
    trois_mois_avant = datetime.now() - timedelta(days=90)
    
    # Filtrer les factures en attente depuis plus de 3 mois
    factures_en_attente_longue = file1_en_attente_df[
        pd.to_datetime(file1_en_attente_df['Date Facture']) < trois_mois_avant
    ]

    # Vérifier s'il y a des factures qui correspondent
    if not factures_en_attente_longue.empty:
        # Contenu de l'alerte
        alerte_texte = f"Attention : {len(factures_en_attente_longue)} facture(s) sont en attente depuis plus de 3 mois.\n\n"
        for index, row in factures_en_attente_longue.iterrows():
            alerte_texte += f"Facture N°: {row['N° Facture']}, Date: {row['Date Facture']}, Montant: {row['Montant']:.2f}\n"

        # Configuration de l'email
        email_from = 'elbakalimalak312@gmail.com'
        email_to = 'elbakali.malak@etu.uae.ac.ma'
        email_subject = 'Alerte : Factures en attente depuis plus de 3 mois'

        # Configuration du message
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = email_to
        msg['Subject'] = email_subject
        msg.attach(MIMEText(alerte_texte, 'plain'))

        # Configuration du serveur SMTP pour Gmail
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        smtp_username = 'elbakalimalak312@gmail.com'
        smtp_password = '*****'  

        # Envoyer l'email
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(email_from, email_to, msg.as_string())
            server.quit()
            print('Alerte pour les factures en attente envoyée avec succès')
        except Exception as e:
            print(f'Erreur lors de l\'envoi de l\'alerte: {e}')
        

# Planifier la vérification tous les jours à 09:00
schedule.every(10).seconds.do(verifier_factures_en_attente)

# Garder le script en cours d'exécution
while True:
    schedule.run_pending()
    time.sleep(1)
