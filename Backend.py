from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

file_name = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qr_data.xlsx")

# Capacités fixes de chaque zone
CAPACITES = {'TS01': 1000, 'TS02': 300, 'TS03': 1700}
ZONES = ['TS01', 'TS02', 'TS03']

def calculate_monthly_rotation(df, target_month=None):
    """Calcule rotation UNIQUEMENT pour le mois en cours"""
    if target_month is None:
        target_month = pd.to_datetime(df['DATE'].iloc[-1]).month
    
    # FILTRER Sheet1 par MOIS COURANT UNIQUEMENT
    df_monthly = df[df['DATE'].apply(lambda x: pd.to_datetime(x).month == target_month)].copy()
    
    rotation_rows = []
    for zone in ZONES:
        palettes_passees = df_monthly[df_monthly['EMPLACEMENT'] == zone]['QR_CODE'].unique()
        sorties = 0
        
        for pal in palettes_passees:
            hist = df_monthly[df_monthly['QR_CODE'] == pal].reset_index(drop=True)
            in_zone_idx = hist[hist['EMPLACEMENT'] == zone].index
            for idx in in_zone_idx:
                if idx + 1 < len(hist) and hist.loc[idx + 1, 'EMPLACEMENT'] != zone:
                    sorties += 1
        
        stock_moyen = CAPACITES.get(zone, 1)
        rotation = round(sorties / stock_moyen, 4) if stock_moyen > 0 else 0
        rotation_rows.append({
            'EMPLACEMENT': zone,
            'NB_SORTIES': sorties,
            'STOCK_MOYEN': stock_moyen,
            'ROTATION': rotation,
            'MOIS_COURANT': target_month
        })
    return pd.DataFrame(rotation_rows)

# 🔥 NOUVEAU : Historique pour TOUS les mois
def calculate_historical_rotations(df):
    """Calcule rotation POUR TOUS les mois historiques"""
    df['DATE'] = pd.to_datetime(df['DATE'])
    months = sorted(df['DATE'].dt.to_period('M').unique())
    
    all_rotations = []
    for month_period in months:
        month_df = df[df['DATE'].dt.to_period('M') == month_period].copy()
        month_num = month_period.month
        
        month_rows = []
        for zone in ZONES:
            palettes = month_df[month_df['EMPLACEMENT'] == zone]['QR_CODE'].unique()
            sorties = 0
            
            for pal in palettes:
                hist = month_df[month_df['QR_CODE'] == pal].reset_index(drop=True)
                in_zone_idx = hist[hist['EMPLACEMENT'] == zone].index
                for idx in in_zone_idx:
                    if idx + 1 < len(hist) and hist.loc[idx + 1, 'EMPLACEMENT'] != zone:
                        sorties += 1
            
            stock_moyen = CAPACITES.get(zone, 1)
            rotation = round(sorties / stock_moyen, 4) if stock_moyen > 0 else 0
            month_rows.append({
                'MOIS': f"{month_period.year}-{month_period.month:02d}",
                'EMPLACEMENT': zone,
                'NB_SORTIES': sorties,
                'ROTATION': rotation
            })
        
        all_rotations.extend(month_rows)
    
    return pd.DataFrame(all_rotations)

def recalculate_sheets():
    """
    Relit Sheet1 et recalcule dynamiquement STOCK, OCCUPATION, ROTATION + HISTORIQUE.
    """
    df = pd.read_excel(file_name, sheet_name='Sheet1')
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=['QR_CODE'])
    df['EMPLACEMENT'] = df['EMPLACEMENT'].fillna('INCONNU')
    df = df[~df['EMPLACEMENT'].isin(['INCONNU', 'UNKNOWN'])]
    df['DATE'] = df['DATE'].astype(str)
    df['HEURE'] = df['HEURE'].astype(str)
    df = df.sort_values(['QR_CODE', 'DATE', 'HEURE'])

    # STOCK : dernière position connue
    df_valid = df[df['EMPLACEMENT'] != 'SORTIE']
    last_pos = df_valid.groupby('QR_CODE').last().reset_index()[['QR_CODE', 'EMPLACEMENT']]
    stock = last_pos.groupby('EMPLACEMENT')['QR_CODE'].count().reset_index()
    stock.columns = ['EMPLACEMENT', 'QUANTITE']

    # OCCUPATION
    occupation_rows = []
    for zone in ZONES:
        qte = int(stock[stock['EMPLACEMENT'] == zone]['QUANTITE'].sum())
        cap = CAPACITES.get(zone, 1)
        taux = round((qte / cap) * 100, 2)
        occupation_rows.append({
            'EMPLACEMENT': zone,
            'CAPACITE': cap,
            'QUANTITE': qte,
            'TAUX_OCCUPATION (%)': taux
        })
    df_occ = pd.DataFrame(occupation_rows)

    # ROTATION MENSUELLE COURANTE (RESET AUTO)
    df_rot = calculate_monthly_rotation(df)

    # ÉCRITURE + HISTORIQUE
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_occ.to_excel(writer, sheet_name='OCCUPATION', index=False)
        df_rot.to_excel(writer, sheet_name='ROTATION', index=False)
        stock.to_excel(writer, sheet_name='STOCK', index=False)
        
        # NOUVELLE FEUILLE HISTORIQUE
        df_histo = calculate_historical_rotations(df)
        df_histo.to_excel(writer, sheet_name='ROTATION_HISTO', index=False)

@app.route("/")
def index():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    return send_from_directory(root_dir, "dashboard.html")

@app.route("/stats")
def stats():
    try:
        df_occ = pd.read_excel(file_name, sheet_name="OCCUPATION")
        df_rot = pd.read_excel(file_name, sheet_name="ROTATION")

        occupation = []
        for _, row in df_occ.iterrows():
            occupation.append({
                "emplacement": str(row['EMPLACEMENT']),
                "quantite": int(row['QUANTITE']),
                "capacite": int(row['CAPACITE']),
                "taux": float(row['TAUX_OCCUPATION (%)'])
            })

        rotation = []
        for _, row in df_rot.iterrows():
            rotation.append({
                "emplacement": str(row['EMPLACEMENT']),
                "nb_sorties": int(row['NB_SORTIES']),
                "stock_moyen": int(row['STOCK_MOYEN']),
                "rotation": float(row['ROTATION'])
            })

        return jsonify({"occupation": occupation, "rotation": rotation})
    except Exception as e:
        print(f"Erreur Stats: {e}")
        return jsonify({"occupation": [], "rotation": []})

@app.route("/stats/historique")
def stats_historique():
    """🔥 API pour historique mensuel (Graphiques !)"""
    try:
        df_histo = pd.read_excel(file_name, sheet_name="ROTATION_HISTO")
        
        histo_data = {}
        for zone in ZONES:
            zone_data = df_histo[df_histo['EMPLACEMENT'] == zone]
            histo_data[zone] = {
                'mois': zone_data['MOIS'].tolist(),
                'rotations': zone_data['ROTATION'].tolist(),
                'sorties': zone_data['NB_SORTIES'].tolist()
            }
        
        return jsonify(histo_data)
    except:
        return jsonify({"TS01": {"mois": [], "rotations": []}, 
                       "TS02": {"mois": [], "rotations": []}, 
                       "TS03": {"mois": [], "rotations": []}})

@app.route("/stats/tendance")
def stats_tendance():
    try:
        today = datetime.now()
        labels = [(today - pd.Timedelta(days=i)).strftime("%d %b") for i in range(6, -1, -1)]

        df = pd.read_excel(file_name, sheet_name="Sheet1")
        df.columns = df.columns.str.strip()
        df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')

        tendance = {}
        for zone in ZONES:
            counts = []
            for i in range(6, -1, -1):
                day = today - pd.Timedelta(days=i)
                count = len(df[(df['EMPLACEMENT'] == zone) & (df['DATE'].dt.date == day.date())])
                counts.append(count)
            tendance[zone] = counts

        return jsonify({"labels": labels, **tendance})
    except Exception as e:
        print(f"Erreur Tendance: {e}")
        today = datetime.now()
        labels = [(today - pd.Timedelta(days=i)).strftime("%d %b") for i in range(6, -1, -1)]
        return jsonify({"labels": labels, "TS01": [0]*7, "TS02": [0]*7, "TS03": [0]*7})

@app.route("/stats/scans")
def stats_scans():
    try:
        df = pd.read_excel(file_name, sheet_name="Sheet1")
        df.columns = df.columns.str.strip()
        df_tail = df.tail(15).copy()
        scans_api = []
        for _, row in df_tail.iterrows():
            scans_api.append({
                "palette_id": str(row.get("QR_CODE", "")),
                "emplacement": str(row.get("EMPLACEMENT", "RECEPTION")),
                "date_scan": str(row.get("DATE", "")),
                "heure_scan": str(row.get("HEURE", ""))
            })
        return jsonify(scans_api[::-1])
    except Exception as e:
        print(f"Erreur Scans: {e}")
        return jsonify([])

@app.route("/scan", methods=["POST"])
def scan():
    try:
        data = request.get_json()
        new_row = {
            "QR_CODE": data["qr"],
            "EMPLACEMENT": data["loc"],
            "DATE": datetime.now().strftime("%Y-%m-%d"),
            "HEURE": datetime.now().strftime("%H:%M:%S")
        }

        # 1. Ajouter le scan dans Sheet1
        df = pd.read_excel(file_name, sheet_name="Sheet1")
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

        # 2. Recalculer TOUT (mensuel + historique)
        recalculate_sheets()

        return jsonify({"status": "ok", "message": "Scan OK + Rotation mensuelle + Historique mis à jour"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True, ssl_context='adhoc')
