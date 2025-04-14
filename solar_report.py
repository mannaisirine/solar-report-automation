import os
import pandas as pd
import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import io
import logging
from datetime import datetime

# Configuration via variables d'environnement
EMAIL = os.environ['EMAIL']
APP_PASSWORD = os.environ['APP_PASSWORD']
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587'))
IMAP_SERVER = os.environ.get('IMAP_SERVER', 'imap.gmail.com')

# Configurer les logs
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def connecter_imap():
    """Connexion sécurisée au serveur IMAP"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, APP_PASSWORD)
        mail.select("INBOX")
        return mail
    except Exception as e:
        logger.error(f"Erreur IMAP: {str(e)}", exc_info=True)
        return None

def telecharger_rapport_solarman():
    """Télécharge le dernier rapport Excel depuis l'email"""
    mail = None
    try:
        mail = connecter_imap()
        if not mail:
            return None
        
        status, messages = mail.search(None, '(FROM "noreply@notice.solarmanpv.com")')
        if status != 'OK':
            logger.warning("Aucun email trouvé de cet expéditeur.")
            return None
        
        message_ids = messages[0].split()
        for email_id in reversed(message_ids):
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            if status != 'OK':
                continue
            
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        
                        filename = part.get_filename()
                        if filename and filename.lower().endswith(('.xlsx', '.xls')):
                            file_content = part.get_payload(decode=True)
                            df = pd.read_excel(io.BytesIO(file_content))
                            logger.info(f"Fichier trouvé: {filename}")
                            return df
        
        logger.warning("Aucun fichier Excel trouvé.")
        return None
        
    except Exception as e:
        logger.error(f"Erreur lors du traitement: {str(e)}", exc_info=True)
        return None
    finally:
        if mail:
            try:
                mail.close()
                mail.logout()
            except:
                pass

def analyser_donnees(df):
    """Analyse les données et génère des alertes"""
    if df.empty:
        logger.error("DataFrame vide reçu")
        return None, None

    # Nettoyage des colonnes
    df.columns = df.columns.str.strip().str.lower().str.replace('\u00a0', ' ').str.replace(' ', '_')
    
    # Détection colonne nom
    colonne_nom = next((col for col in df.columns if 'nom' in col or 'centrale' in col or 'plant' in col), None)
    if not colonne_nom:
        logger.error("Colonne 'nom_centrale' introuvable")
        return None, None
    
    df = df.rename(columns={colonne_nom: 'nom_centrale'})
    
    # Conversion numérique
    try:
        df['total'] = pd.to_numeric(df['total'], errors='coerce').fillna(0)
    except KeyError:
        logger.error("Colonne 'total' manquante")
        return None, None

    # Détection des régions (exemple simplifié)
    df['region'] = df.get('region', 'Inconnue')
    
    # Détection installations nulles
    installations_nulle = df[df["total"] == 0].copy()
    
    # Analyse des performances
    regions_data = []
    for region in df[df["total"] > 0]["region"].unique():
        region_df = df[df["region"] == region]
        moyenne = region_df["total"].mean()
        seuil = 0.9 * moyenne
        
        sous_perf = region_df[(region_df["total"] > 0) & (region_df["total"] < seuil)]
        regions_data.append({
            "region": region,
            "moyenne": moyenne,
            "seuil": seuil,
            "sous_performantes": sous_perf
        })
    
    return installations_nulle, regions_data

def envoyer_email(destinataire, sujet, installations_nulle, regions_data):
    """Envoi du rapport par email"""
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL
        msg['To'] = destinataire
        msg['Subject'] = sujet
        
        # Version texte
        text_content = f"""Rapport du {datetime.now().strftime('%d/%m/%Y')}
        
Installations hors service: {len(installations_nulle)}
{installations_nulle[['nom_centrale', 'region']].to_string()}

Détails par région:
"""
        for region in regions_data:
            text_content += f"\n{region['region']} - Moyenne: {region['moyenne']:.2f} kWh\n"
            text_content += region['sous_performantes'][['nom_centrale', 'total']].to_string()
        
        msg.attach(MIMEText(text_content, 'plain'))
        
        # Envoi
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, APP_PASSWORD)
            server.send_message(msg)
        
        logger.info("Email envoyé avec succès")
        return True
    except Exception as e:
        logger.error(f"Erreur d'envoi: {str(e)}", exc_info=True)
        return False

if __name__ == "__main__":
    logger.info("Début de l'exécution")
    
    # Téléchargement et analyse
    df = telecharger_rapport_solarman()
    if df is not None:
        installations_nulle, regions_data = analyser_donnees(df)
        if installations_nulle is not None:
            # Envoi du rapport
            envoyer_email(
                destinataire=EMAIL,  # Envoyer à soi-même
                sujet=f"Rapport photovoltaïque {datetime.now().strftime('%d/%m')}",
                installations_nulle=installations_nulle,
                regions_data=regions_data
            )
    
    logger.info("Exécution terminée")