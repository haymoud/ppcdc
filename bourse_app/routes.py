from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import re
from io import BytesIO

# Création du Blueprint 'bourse'
bourse_bp = Blueprint('bourse', __name__, template_folder='templates')

# =========================
# 1. COBOL PARSER
# =========================
def parse_cobol(file_bytes):
    text = file_bytes.decode("latin-1", errors="ignore")
    lines = text.splitlines()

    net_pattern = re.compile(r"\d{4,7},\d{2}")
    dev_pattern = re.compile(r"(\d+)\s*Euro", re.IGNORECASE)

    rows = []

    for line in lines:
        if not line.startswith("I "):
            continue

        tokens = line.split()

        # ORDRE
        ordre = next((t for t in tokens if re.fullmatch(r"\d{5}", t)), None)
        if not ordre:
            continue

        # COMPTE
        compte = next((t for t in tokens if re.fullmatch(r"\d{8,12}", t)), None)
        if not compte:
            continue

        # NET
        net_match = net_pattern.search(line)
        if not net_match:
            continue
        net = net_match.group()

        # DEVISE
        dev_match = dev_pattern.search(line)
        devise = dev_match.group(1) if dev_match else ""

        # NAME
        try:
            name = line[line.index(compte)+len(compte):line.index(net)].strip()
        except:
            name = ""

        rows.append([ordre, compte, name, net, devise])

    df = pd.DataFrame(rows, columns=["ORDRE", "COMPTE", "NAME", "NET", "DEVISE"])

    # Insert blank line before new sequence
    out = []
    for i, row in df.iterrows():
        if row["ORDRE"] == "00001" and i != 0:
            out.append({c: "" for c in df.columns})
        out.append(row.to_dict())

    return pd.DataFrame(out)


# =========================
# 2. COMPARISON ENGINE
# =========================
def standardize(df, label):
    # Trouver les colonnes dynamiquement
    try:
        nni = [c for c in df.columns if "NNI" in str(c).upper()][0]
        name = [c for c in df.columns if "NOM" in str(c).upper() or "NAME" in str(c).upper()][0]
        net = [c for c in df.columns if "NET" in str(c).upper()][0]
    except IndexError:
        raise ValueError("Les colonnes requises (NNI, NOM/NAME, NET) sont introuvables.")

    df = df[[nni, name, net]].copy()
    df.columns = ["NNI", f"Name_{label}", f"Net_{label}"]

    df["NNI"] = df["NNI"].astype(str).str.replace(" ", "").str.zfill(10)
    df[f"Net_{label}"] = pd.to_numeric(df[f"Net_{label}"], errors="coerce").fillna(0)

    return df.groupby("NNI").agg({
        f"Name_{label}": "first",
        f"Net_{label}": "sum"
    }).reset_index()


# =========================
# ROUTES DU BLUEPRINT
# =========================

@bourse_bp.route("/")
def home():
    return render_template("bourse_index.html")

@bourse_bp.route("/convert", methods=["POST"])
def convert():
    if 'file' not in request.files or request.files['file'].filename == '':
        flash("Aucun fichier COBOL sélectionné pour la conversion.")
        return redirect('/bourse')

    try:
        file = request.files["file"]
        df = parse_cobol(file.read())
        
        if df.empty:
            flash("Aucune donnée valide trouvée dans le fichier COBOL.")
            return redirect('/bourse')

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="cobol_converti.xlsx")
    except Exception as e:
        flash(f"Erreur lors de la conversion : {str(e)}")
        return redirect('/bourse')


@bourse_bp.route("/compare", methods=["POST"])
def compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        flash("Veuillez sélectionner les deux fichiers Excel pour la comparaison.")
        return redirect('/bourse')

    f1 = request.files["file1"]
    f2 = request.files["file2"]

    if f1.filename == '' or f2.filename == '':
        flash("Les deux fichiers Excel sont requis.")
        return redirect('/bourse')

    try:
        df1 = pd.read_excel(f1)
        df2 = pd.read_excel(f2)

        df1 = standardize(df1, "A")
        df2 = standardize(df2, "B")

        merged = pd.merge(df1, df2, on="NNI", how="outer").fillna(0)
        merged["Difference"] = merged["Net_A"] - merged["Net_B"]

        def obs(r):
            if r["Difference"] == 0:
                return "Match"
            if r["Net_A"] == 0:
                return "Missing A"
            if r["Net_B"] == 0:
                return "Missing B"
            return "Mismatch"

        merged["Observation"] = merged.apply(obs, axis=1)

        output = BytesIO()
        merged.to_excel(output, index=False)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="comparaison_bourse.xlsx")
    
    except ValueError as ve:
        flash(str(ve))
        return redirect('/bourse')
    except Exception as e:
        flash(f"Erreur lors de la comparaison : {str(e)}")
        return redirect('/bourse')