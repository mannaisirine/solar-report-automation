import os
import pandas as pd
import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import io
import traceback
from datetime import datetime
import logging

# Configuration via variables d'environnement
EMAIL = os.environ['EMAIL']
APP_PASSWORD = os.environ['APP_PASSWORD']
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587'))
IMAP_SERVER = os.environ.get('IMAP_SERVER', 'imap.gmail.com')
DEST=os.environ['DEST']

# Configuration des logs
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
        logger.error(f"Erreur de connexion IMAP: {str(e)}")
        logger.error(traceback.format_exc())
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
        if not message_ids:
            logger.warning("Aucun email trouvé dans la boîte de réception.")
            return None
        
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
        
        logger.warning("Aucun fichier Excel trouvé dans les pièces jointes.")
        return None
        
    except Exception as e:
        logger.error(f"Erreur lors du traitement: {str(e)}")
        logger.error(traceback.format_exc())
        return None
    finally:
        if mail:
            try:
                mail.close()
                mail.logout()
            except:
                pass

def analyser_donnees(df):
    """Analyse les données avec la même logique que la version Streamlit"""
    if df.empty:
        return None, None

    # Nettoyage des colonnes
    df.columns = df.columns.str.strip().str.lower().str.replace('\u00a0', ' ').str.replace(' ', '_')
    
    # Identification colonne nom
    colonne_nom = next((col for col in df.columns if 'nom' in col or 'centrale' in col or 'plant' in col), None)
    if not colonne_nom:
        logger.error("Colonne 'nom_centrale' introuvable.")
        return None, None
    
    df = df.rename(columns={colonne_nom: 'nom_centrale'})
    
    if 'total' not in df.columns:
        logger.error(f"Colonne 'total' introuvable. Colonnes disponibles: {df.columns.tolist()}")
        return None, None

    df['total'] = pd.to_numeric(df['total'], errors='coerce').fillna(0)
    
    # Gestion région
    if 'region' not in df.columns:
        for col in df.columns:
            if 'region' in col.lower() or 'région' in col.lower():
                df = df.rename(columns={col: 'region'})
                break
        else:
            df['region'] = 'Inconnue'
    
    df['region'] = df['region'].astype(str)

    def detect_region_personnalisee(region_str):
        if pd.isna(region_str):
            return "Autres"
        region_str = region_str.strip().lower()
        if region_str.startswith("tunisie/sfax"):
            return "Sfax"
        elif region_str in ["tunisie", "tunisie/ariana", "tunisie/tunis", "tunisie/ariana/ariana médina", "tunisie/ariana/soukra"]:
            return "Tunis"
        elif "médenine" in region_str:
            return "Médenine"
        elif "gabès" in region_str:
            return "Gabès"
        elif "nabeul" in region_str:
            return "Nabeul"
        elif "monastir" in region_str:
            return "Monastir"
        elif "sousse" in region_str:
            return "Sousse"
        elif "bizerte" in region_str:
            return "Bizerte"
        elif "mahdia" in region_str:
            return "Mahdia"
        else:
            return "Autres"

    df["region_perso"] = df["region"].apply(detect_region_personnalisee)

    # Installations nulles
    installations_nulle = df[df["total"] == 0].copy()
    
    # Régions avec production > 0
    regions_avec_production = df[df["total"] > 0]["region_perso"].unique()
    
    regions_data = []
    for region in regions_avec_production:
        region_df = df[df["region_perso"] == region]
        if len(region_df) == 0:
            continue
            
        moyenne_region = region_df[region_df["total"] > 0]["total"].mean()
        seuil_performance = 0.9 * moyenne_region
        max_region = region_df["total"].max()
        
        sous_performantes = region_df[(region_df["total"] > 0) & (region_df["total"] < seuil_performance)]
        
        regions_data.append({
            "region": region,
            "moyenne": moyenne_region,
            "seuil_performance": seuil_performance,
            "max": max_region,
            "sous_performantes": sous_performantes,
            "nb_sous_performantes": len(sous_performantes)
        })
    
    return installations_nulle, regions_data

def envoyer_email(destinataire, sujet, installations_nulle, regions_data):
    """Version identique à la fonction Streamlit originale"""
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL
        msg['To'] = destinataire
        msg['Subject'] = sujet
        
        html_content = f"""
        <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    h1 {{ color: #2c3e50; }}
                    h2 {{ color: #3498db; border-bottom: 1px solid #eee; padding-bottom: 5px; }}
                    h3 {{ color: #2c3e50; margin-top: 20px; }}
                    .count-badge {{
                        display: inline-block;
                        padding: 3px 8px;
                        background-color: #ff9800;
                        color: white;
                        border-radius: 12px;
                        font-weight: bold;
                        margin-left: 10px;
                    }}
                    table {{
                        border-collapse: collapse;
                        width: 100%;
                        margin: 15px 0;
                    }}
                    th, td {{
                        padding: 10px;
                        text-align: left;
                        border-bottom: 1px solid #ddd;
                    }}
                    th {{
                        background-color: #f2f2f2;
                    }}
                    .alert {{
                        padding: 12px;
                        margin: 15px 0;
                        background: #fff3cd;
                        border-left: 5px solid #ffc107;
                    }}
                    .section {{
                        margin-bottom: 30px;
                    }}
                </style>
            </head>
            <body>
                <h1>Rapport d'Alertes des Centrales Photovoltaïques</h1>
        """
        
        if not installations_nulle.empty:
            html_content += f"""
            <div class="section">
                <h2>🚨 Installations Hors Service (Production = 0) <span class="count-badge">{len(installations_nulle)}</span></h2>
                <table>
                    <tr>
                        <th>Centrale</th>
                        <th>Région</th>
                    </tr>
            """
            
            for _, row in installations_nulle.iterrows():
                html_content += f"""
                <tr>
                    <td>{row['nom_centrale']}</td>
                    <td>{row['region_perso']}</td>
                </tr>
                """
            
            html_content += """
                </table>
                <div class="alert">
                    <strong>Action recommandée :</strong>  Vérifier les onduleurs et la connexion internet.
                </div>
            </div>
            """
        else:
            html_content += """
            <div class="section">
                <h2>✅ Aucune installation hors service détectée</h2>
            </div>
            """
        
        if regions_data:
            html_content += """
            <div class="section">
                <h2>📊 Analyse des Performances par Région</h2>
            """
            
            for region_info in regions_data:
                region = region_info["region"]
                moyenne = region_info["moyenne"]
                seuil = region_info["seuil_performance"]
                max_power = region_info["max"]
                sous_perf = region_info["sous_performantes"]
                nb_sous_perf = region_info["nb_sous_performantes"]
                
                if nb_sous_perf > 0:
                    html_content += f"""
                    <h3>{region} <span class="count-badge">{nb_sous_perf} sous le seuil</span></h3>
                    <p>Moyenne: {moyenne:.2f} kWh | Seuil (90%): {seuil:.2f} kWh | Max: {max_power:.2f} kWh</p>
                    <table>
                        <tr>
                            <th>Centrale</th>
                            <th>Production (kWh)</th>
                            <th>Écart au seuil</th>
                        </tr>
                    """
                    
                    for _, row in sous_perf.iterrows():
                        ecart = row["total"] - seuil
                        html_content += f"""
                        <tr>
                            <td>{row['nom_centrale']}</td>
                            <td>{row['total']:.2f}</td>
                            <td>{ecart:.2f} kWh</td>
                        </tr>
                        """
                    
                    html_content += """
                    </table>
                    <div class="alert">
                        <strong>Action recommandée :</strong> 
                        <ul>
                            <li>Vérifier l'ombrage</li>
                            <li>Nettoyer les panneaux</li>
                            <li>Contrôler les onduleurs</li>
                        </ul>
                    </div>
                    """
                else:
                    html_content += f"""
                    <h3>{region}</h3>
                    <p>Toutes les centrales sont au-dessus du seuil (Moyenne: {moyenne:.2f} kWh | Seuil: {seuil:.2f} kWh | Max: {max_power:.2f} kWh)</p>
                    """
            
            html_content += """
            </div>
            """
        
        html_content += """
            </body>
        </html>
        """
        
        text_content = "Rapport d'Alertes des Centrales Photovoltaïques\n\n"
        
        if not installations_nulle.empty:
            text_content += f"🚨 Installations Hors Service ({len(installations_nulle)}):\n"
            for _, row in installations_nulle.iterrows():
                text_content += f"- {row['nom_centrale']} ({row['region_perso']})\n"
        
        if regions_data:
            text_content += "\n📊 Performances par Région (seuil = 90% de la moyenne):\n"
            for region_info in regions_data:
                text_content += f"\nRégion {region_info['region']}:\n"
                text_content += f"Moyenne: {region_info['moyenne']:.2f} kWh | Seuil: {region_info['seuil_performance']:.2f} kWh\n"
                if region_info["nb_sous_performantes"] > 0:
                    text_content += f"{region_info['nb_sous_performantes']} sous le seuil:\n"
                    for _, row in region_info["sous_performantes"].iterrows():
                        text_content += f"- {row['nom_centrale']}: {row['total']:.2f} kWh (écart: {row['total'] - region_info['seuil_performance']:.2f} kWh)\n"
        
        msg.attach(MIMEText(text_content, 'plain'))
        msg.attach(MIMEText(html_content, 'html'))
        
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, APP_PASSWORD)
            server.send_message(msg)
        
        logger.info("Email envoyé avec succès!")
        return True
    except Exception as e:
        logger.error(f"Erreur d'envoi: {str(e)}")
        return False

def main():
    logger.info("Début de l'exécution automatique")
    df = telecharger_rapport_solarman()
    if df is not None:
        installations_nulle, regions_data = analyser_donnees(df)
        if installations_nulle is not None:
            envoyer_email(
                DEST,  # Envoyer à vous-même
                f"Rapport photovoltaïque {datetime.now().strftime('%d/%m/%Y')}",
                installations_nulle,
                regions_data
            )
    logger.info("Exécution terminée")

if __name__ == "__main__":
    main()
