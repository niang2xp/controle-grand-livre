# -*- coding: utf-8 -*-
"""
app.py — Contrôle et extraction de grands livres
================================================

Lancement :
    pip install flask pandas openpyxl
    python app.py

Le serveur démarre sur http://127.0.0.1:5000 et ouvre le navigateur.

Fonctionnement :
  1. dépôt d'un ou plusieurs fichiers (.xlsx, .xlsm, .xls, .csv, .txt, .tsv),
     sans plafond de taille dans l'application ;
  2. conversion à l'arrivée : un classeur Excel est converti une fois pour
     toutes en CSV dans un dossier de travail temporaire, effacé à l'arrêt ;
  3. correspondance semi-automatique : l'application propose une colonne pour
     chaque champ attendu, l'utilisateur confirme ou change en voyant un
     aperçu des valeurs de chaque colonne ;
  4. un montant unique est ramené au débit / crédit, par le signe (positif au
     débit par défaut, un bouton + / − inverse) ou par une colonne de sens D / C ;
  5. devise : codes ISO numériques traduits (952 -> XOF, 840 -> USD...),
     XOF par défaut en l'absence de colonne ;
  6. analyse sur un, plusieurs ou tous les fichiers chargés ;
  7. extraction filtrée, colonnes au choix, toutes les lignes consultables
     par défilement, export Excel ou CSV.
"""

import atexit
import glob
import hashlib
import io
import math
import os
import re
import shutil
import tempfile
import threading
import unicodedata
import uuid
import warnings
import webbrowser
from datetime import date, datetime

import numpy as np
import pandas as pd
from flask import Flask, Response, jsonify, request, send_file

app = Flask(__name__)
# Aucun plafond : la vraie limite est la mémoire de la machine.
app.config["MAX_CONTENT_LENGTH"] = None

PREFIXE_TRAVAIL = "grand_livre_"


def preparer_dossier_travail():
    """Dossier temporaire des fichiers déposés, effacé à l'arrêt.

    Les dossiers laissés par une exécution précédente interrompue
    brutalement sont supprimés au démarrage : ils contiennent des données
    client qui n'ont pas à traîner sur le disque.
    """
    for ancien in glob.glob(os.path.join(tempfile.gettempdir(), PREFIXE_TRAVAIL + "*")):
        shutil.rmtree(ancien, ignore_errors=True)
    dossier = tempfile.mkdtemp(prefix=PREFIXE_TRAVAIL)
    atexit.register(shutil.rmtree, dossier, ignore_errors=True)
    return dossier


DOSSIER_TRAVAIL = preparer_dossier_travail()

FICHIERS = {}      # identifiant -> fichier chargé
STRUCTURES = {}    # signature des colonnes -> correspondance partagée
RESULTAT = {"cle": None, "table": None}
VERROU = threading.Lock()


# =====================================================================
#  1. Outils de conversion
# =====================================================================

def normaliser(nom):
    """Forme comparable d'un nom : 'N° pièces ' -> 'n pieces'."""
    texte = unicodedata.normalize("NFKD", str(nom))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    texte = re.sub(r"[^a-z0-9]+", " ", texte.lower())
    return re.sub(r"\s+", " ", texte).strip()


def par_valeurs_uniques(serie, fonction):
    """Applique une conversion aux seules valeurs distinctes d'une colonne
    catégorielle : 45 dates à lire au lieu de 416 000."""
    if not isinstance(serie.dtype, pd.CategoricalDtype):
        return fonction(serie)
    categories = pd.Series(serie.cat.categories, dtype=object)
    if len(categories) == 0:
        return fonction(serie.astype(object))
    codes = serie.cat.codes.to_numpy()
    valeurs = fonction(categories).reset_index(drop=True)
    resultat = valeurs.iloc[np.where(codes >= 0, codes, 0)]
    resultat.index = serie.index
    return resultat.where(pd.Series(codes >= 0, index=serie.index))


def vers_nombre(serie):
    """Nombre lu à la française : '-30000,000' et '1 234,56' compris."""
    if isinstance(serie.dtype, pd.CategoricalDtype):
        return par_valeurs_uniques(serie, vers_nombre).astype("float64")
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    texte = serie.astype("string").str.strip()
    texte = texte.str.replace(r"[\s\u00a0\u202f]", "", regex=True)
    virgule = texte.str.contains(",", na=False)
    texte = texte.mask(virgule, texte.str.replace(".", "", regex=False)
                                     .str.replace(",", ".", regex=False))
    return pd.to_numeric(texte, errors="coerce").astype("float64")


def vers_date(serie):
    """Date lue jour en premier : 01/02/2022 est un 1er février."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    if isinstance(serie.dtype, pd.CategoricalDtype):
        return pd.to_datetime(par_valeurs_uniques(serie, vers_date))
    texte = serie.astype("string").str.strip()
    resultat = pd.to_datetime(texte, errors="coerce", format="%d/%m/%Y")
    for options in ({"format": "ISO8601"}, {"format": "%d/%m/%Y %H:%M:%S"},
                    {"format": "%d/%m/%y"}, {"dayfirst": True}):
        manque = resultat.isna() & texte.notna() & (texte != "")
        if not bool(manque.any()):
            break
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            essai = pd.to_datetime(texte[manque], errors="coerce", **options)
        resultat = resultat.combine_first(essai)
    return resultat


def identifiant(serie):
    """Numéro de compte ou de pièce : texte, sans '.0' ni séparateur."""
    if isinstance(serie.dtype, pd.CategoricalDtype):
        return par_valeurs_uniques(serie, identifiant).astype(object)
    texte = serie.astype("string").str.strip()
    return texte.str.replace(r"^(-?\d+)\.0+$", r"\1", regex=True)


# Codes ISO 4217 numériques les plus probables dans un grand livre
# bancaire d'Afrique de l'Ouest, et au-delà.
ISO_NUMERIQUE = {
    "952": "XOF", "950": "XAF", "840": "USD", "978": "EUR", "826": "GBP",
    "756": "CHF", "124": "CAD", "392": "JPY", "156": "CNY", "682": "SAR",
    "784": "AED", "504": "MAD", "788": "TND", "012": "DZD", "818": "EGP",
    "566": "NGN", "936": "GHS", "324": "GNF", "270": "GMD", "132": "CVE",
    "929": "MRU", "430": "LRD", "404": "KES", "710": "ZAR", "356": "INR",
    "036": "AUD", "578": "NOK", "752": "SEK", "208": "DKK", "414": "KWD",
    "634": "QAR", "048": "BHD", "512": "OMR", "400": "JOD", "344": "HKD",
    "702": "SGD", "643": "RUB", "949": "TRY", "986": "BRL", "646": "RWF",
    "800": "UGX", "834": "TZS", "230": "ETB", "973": "AOA", "976": "CDF",
    "480": "MUR",
}
ALIAS_DEVISES = {
    "FCFA": "XOF", "F CFA": "XOF", "F.CFA": "XOF", "CFA": "XOF",
    "EURO": "EUR", "EUROS": "EUR", "€": "EUR",
    "DOLLAR": "USD", "DOLLARS": "USD", "US$": "USD", "$": "USD",
}
DEVISES_CONNUES = set(ISO_NUMERIQUE.values()) | set(ALIAS_DEVISES.values())


def code_devise(valeur):
    texte = str(valeur).strip().upper()
    if not texte or texte in ("NAN", "<NA>", "NONE"):
        return None
    if re.fullmatch(r"\d+(\.0+)?", texte):
        code = texte.split(".")[0].zfill(3)
        return ISO_NUMERIQUE.get(code, code)
    return ALIAS_DEVISES.get(texte, texte)


def serie_devise(serie):
    """Devise en code alphabétique. Cellule vide : XOF."""
    table = {v: code_devise(v) for v in pd.unique(serie.dropna())}
    return serie.astype(object).map(table).fillna("XOF").astype(object)


def ressemble_devise(serie):
    """Au moins 80 % des valeurs renseignées sont des devises connues."""
    valeurs = serie.dropna()
    if len(valeurs) == 0:
        return False
    comptes = valeurs.value_counts()
    connues = sum(n for v, n in comptes.items() if code_devise(v) in DEVISES_CONNUES)
    return connues / len(valeurs) >= 0.8


MOTS_DEBIT = {"d", "db", "dt", "debit", "debiteur"}
MOTS_CREDIT = {"c", "cr", "ct", "credit", "crediteur"}


def serie_sens(serie):
    """Série 'D' / 'C' tirée d'une colonne de sens, ou None si inexploitable."""
    table = {}
    for v in pd.unique(serie.dropna()):
        mot = normaliser(v)
        table[v] = "D" if mot in MOTS_DEBIT else ("C" if mot in MOTS_CREDIT else None)
    sens = serie.astype(object).map(table)
    if float(sens.notna().mean()) < 0.5:
        return None
    return sens


# =====================================================================
#  2. Détection proposée des colonnes
# =====================================================================

LIBELLES = {
    "compte": "Compte", "debit": "Débit", "credit": "Crédit", "montant": "Montant",
    "sens": "Sens", "piece": "Pièce", "libelle": "Libellé", "impute": "Imputé par",
    "autorise": "Autorisé par", "devise": "Devise", "date": "Date de transaction",
    "date_valeur": "Date de valeur",
}
CLES = list(LIBELLES)

# Les synonymes ne font que proposer : chaque choix se confirme ou se
# corrige dans l'interface, au vu des valeurs de la colonne.
# CREATED_BY n'est plus proposé pour Imputé par : créer une ligne dans le
# système n'est pas l'imputer.
SYNONYMES = {
    "compte": {
        "exact": ["n compte", "no compte", "num compte", "numero compte", "numero de compte",
                  "compte general", "compte", "cpte", "account", "account number", "gl account",
                  "gl code", "code gl", "gl", "compte gl", "account code", "code compte",
                  "chapitre", "chapitre comptable"],
        "contient": ["n compte", "numero compte", "compte", "cpte", "account", "gl code"],
    },
    "debit": {"exact": ["debit", "montant debit", "dt", "mouvement debit"],
              "contient": ["debit"]},
    "credit": {"exact": ["credit", "montant credit", "ct", "mouvement credit"],
               "contient": ["credit"]},
    "montant": {"exact": ["montant", "montant signe", "amount", "valeur", "value", "mnt",
                          "cv amount", "montant cv", "contre valeur", "mon"],
                "contient": ["montant", "amount"]},
    "sens": {"exact": ["sens", "sens ecriture", "sens de l ecriture", "sens du montant",
                       "d c", "dc", "debit credit", "code sens", "sens operation", "sen"],
             "contient": ["sens"]},
    "piece": {
        "exact": ["n piece", "n pieces", "no piece", "num piece", "numero piece",
                  "numero de piece", "piece", "pieces", "document", "n document", "voucher",
                  "op no", "no op", "num op", "numero op", "op number", "no operation",
                  "numero operation", "transaction id", "trans no", "pie"],
        "contient": ["piece", "voucher", "document", "op no"],
    },
    "libelle": {"exact": ["libelle", "libelles", "intitule", "designation", "description",
                          "objet", "narration", "text", "lib", "label"],
                "contient": ["libelle", "intitule", "designation", "description", "narration"]},
    "impute": {"exact": ["impute", "impute par", "imputeur", "imputed by"],
               "contient": ["impute par", "imputeur", "imputed by"]},
    "autorise": {"exact": ["autorise", "autorise par", "valide par", "approuve par",
                           "approbateur", "validateur", "approved by", "authorized by",
                           "authorised by"],
                 "contient": ["autorise", "authoris", "authoriz", "valide par", "validateur",
                              "approuve par", "approved by", "approbateur"]},
    "devise": {"exact": ["devise", "monnaie", "currency", "currency code", "code devise",
                         "ccy", "code monnaie", "dev"],
               "contient": ["devise", "currency", "monnaie"]},
    "date": {"exact": ["date", "date comptable", "date ecriture", "date piece", "date operation",
                       "posting date", "trans date", "transaction date", "date transaction",
                       "date de transaction", "dco"],
             "contient": ["date"]},
    # « date » seul figure dans la date de valeur : on exige les deux mots,
    # sinon « Montant contre valeur » serait pris pour une date.
    "date_valeur": {"exact": ["date valeur", "date de valeur", "value date", "valeur date",
                              "date val", "dt valeur", "dva"],
                    "contient": ["date valeur", "date de valeur", "value date"]},
}

ORDRE_DETECTION = ["debit", "credit", "sens", "devise", "compte", "piece", "libelle",
                   "impute", "autorise", "date_valeur", "date", "montant"]


def score_colonne(df, col, cle):
    """Valeurs exploitables : départage FC_AMOUNT (presque vide) et CV_AMOUNT."""
    if cle in ("montant", "debit", "credit"):
        valeurs = vers_nombre(df[col])
        return int((valeurs.notna() & (valeurs != 0)).sum())
    return int(df[col].notna().sum())


def detecter_colonnes(df):
    colonnes = list(df.columns)
    formes = {col: normaliser(col) for col in colonnes}
    prises, mapping = set(), {}

    for cle in ORDRE_DETECTION:
        for col in colonnes:
            if col not in prises and formes[col] in SYNONYMES[cle]["exact"]:
                prises.add(col)
                mapping[cle] = col
                break

    for cle in ORDRE_DETECTION:
        if cle in mapping:
            continue
        candidats = [col for col in colonnes if col not in prises
                     and any(f in formes[col] for f in SYNONYMES[cle]["contient"])]
        if not candidats:
            continue
        candidats.sort(key=lambda col: -score_colonne(df, col, cle))
        prises.add(candidats[0])
        mapping[cle] = candidats[0]

    # La proposition de devise et de sens se vérifie sur les valeurs.
    if "devise" in mapping and not ressemble_devise(df[mapping["devise"]]):
        del mapping["devise"]
    if "sens" in mapping and serie_sens(df[mapping["sens"]]) is None:
        del mapping["sens"]
    return mapping


def nouvelle_structure(df):
    proposition = detecter_colonnes(df)
    a_dc = "debit" in proposition and "credit" in proposition
    mode = "montant" if ("montant" in proposition and not a_dc) else "dc"
    return {"mapping": {cle: proposition.get(cle) for cle in CLES},
            "mode": mode, "inverse": False, "version": 1,
            "colonnes": [str(c) for c in df.columns]}


def signature(colonnes):
    return hashlib.sha1("\x1f".join(map(str, colonnes)).encode("utf-8")).hexdigest()[:16]


# =====================================================================
#  3. Lecture et conversion des fichiers
# =====================================================================

EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".tsv")


def compacter(df):
    """Mémoire divisée par dix sur un historique bancaire : une colonne dont
    les valeurs se répètent les stocke une seule fois. Une cellule faite
    d'espaces seulement devient vide, sinon elle ne compterait pas comme
    valeur manquante."""
    for col in df.columns:
        serie = df[col]
        if serie.nunique(dropna=True) < 0.5 * max(len(serie), 1):
            serie = serie.astype("category")
            blancs = [c for c in serie.cat.categories if not str(c).strip()]
            if blancs:
                serie = serie.cat.remove_categories(blancs)
        else:
            serie = serie.mask(serie.str.strip().eq("").fillna(False))
        df[col] = serie
    return df


def lire_csv(chemin):
    """Fichier délimité lu en texte, chaque valeur telle qu'écrite.

    Retourne (table, coupure). Aucune ligne n'est retirée : si le fichier
    s'arrête au milieu d'un enregistrement, cette dernière ligne est gardée
    avec ce qu'elle contient, et coupure décrit ce qui lui manque.
    """
    with open(chemin, "rb") as flux:
        echantillon = flux.read(256 * 1024)
        flux.seek(0, os.SEEK_END)
        taille = flux.tell()
        flux.seek(max(0, taille - 65536))
        fin = flux.read()
    for encodage in ("utf-8-sig", "cp1252"):
        try:
            extrait = echantillon.decode(encodage)
            break
        except UnicodeDecodeError:
            continue
    else:
        encodage, extrait = "latin-1", echantillon.decode("latin-1")
    premiere = extrait.splitlines()[0] if extrait else ""
    separateur = max([";", ",", "\t", "|"], key=premiere.count)
    # Certains exports terminent chaque ligne par un séparateur de trop.
    # Sans index_col=False, pandas prend la première colonne pour un index
    # et décale toutes les autres d'un cran, sans le signaler.
    options = dict(sep=separateur, dtype=str, engine="c", low_memory=False, index_col=False)
    df = None
    for essai in dict.fromkeys((encodage, "cp1252", "latin-1")):
        try:
            df = pd.read_csv(chemin, encoding=essai, **options)
            break
        except UnicodeDecodeError:
            continue
    derniere = fin.rstrip(b"\r\n").rsplit(b"\n", 1)[-1]
    champs = derniere.count(separateur.encode()) + 1
    coupure = None
    if len(df) > 0 and not fin.endswith(b"\n") and champs < len(df.columns):
        coupure = {"champs": champs, "attendus": len(df.columns)}
    return compacter(df), coupure


def convertir_excel(chemin_source, chemin_csv, feuille=None):
    """Classeur -> CSV. Retourne (feuilles, feuille convertie)."""
    ext = os.path.splitext(chemin_source)[1].lower()
    moteur = "xlrd" if ext == ".xls" else "openpyxl"
    try:
        classeur = pd.ExcelFile(chemin_source, engine=moteur)
    except ImportError:
        raise ValueError("Les fichiers .xls demandent la bibliothèque xlrd : "
                         "« pip install xlrd », ou enregistrez le fichier en .xlsx.")
    feuilles = [str(f) for f in classeur.sheet_names]
    cible = feuille if feuille in feuilles else feuilles[0]
    df = classeur.parse(cible)
    # Les valeurs sont écrites comme on les lirait dans Excel : dates au
    # format jour/mois/année, entiers sans « .0 ».
    for col in df.columns:
        serie = df[col]
        if pd.api.types.is_datetime64_any_dtype(serie):
            valeurs = serie.dropna()
            minuit = bool(len(valeurs) == 0 or (valeurs == valeurs.dt.normalize()).all())
            df[col] = serie.dt.strftime("%d/%m/%Y" if minuit else "%d/%m/%Y %H:%M:%S")
        elif pd.api.types.is_float_dtype(serie):
            valeurs = serie.dropna()
            if len(valeurs) and bool((valeurs % 1 == 0).all()):
                df[col] = serie.astype("Int64")
    df.to_csv(chemin_csv, sep=";", index=False, encoding="utf-8")
    return feuilles, cible


def charger(fichier):
    """Lit le CSV de travail et rattache le fichier à sa structure."""
    df, coupure = lire_csv(fichier["csv"])
    if df.empty:
        raise ValueError("Le fichier ne contient aucune ligne.")
    sig = signature(df.columns)
    if sig not in STRUCTURES:
        STRUCTURES[sig] = nouvelle_structure(df)
    # Aperçu des colonnes calculé une fois : il ne dépend pas des choix de
    # correspondance, inutile de le refaire à chaque analyse.
    apercus = {str(col): apercu(df[col]) for col in df.columns}
    fichier.update({"df": df, "signature": sig, "lignes": int(len(df)), "cache": None,
                    "coupure": coupure, "apercus": apercus})


def infos_fichier(f):
    return {"id": f["id"], "nom": f["nom"], "lignes": f["lignes"], "taille": f["taille"],
            "feuilles": f["feuilles"], "feuille": f["feuille"], "signature": f["signature"]}


# =====================================================================
#  4. Table normalisée
# =====================================================================

COLONNES_TABLE = ["Fichier", "Compte", "Débit", "Crédit", "Pièce", "Libellé", "Imputé par",
                  "Autorisé par", "Date de transaction", "Date de valeur", "Devise"]


def montants(df, structure):
    """(débit, crédit) selon la forme des montants choisie pour la structure."""
    m, vide = structure["mapping"], pd.Series(np.nan, index=df.index)
    if structure["mode"] == "dc":
        debit = vers_nombre(df[m["debit"]]) if m.get("debit") else None
        credit = vers_nombre(df[m["credit"]]) if m.get("credit") else None
        return debit, credit
    if not m.get("montant"):
        return None, None
    montant = vers_nombre(df[m["montant"]])
    sens = serie_sens(df[m["sens"]]) if m.get("sens") else None
    if sens is not None:
        debit = montant.abs().where(sens == "D", vide)
        credit = montant.abs().where(sens == "C", vide)
    else:
        # Par défaut, positif au débit et négatif au crédit : c'est la règle
        # vérifiée sur les douze GL 2022. Le bouton + / − l'inverse.
        debit = montant.where(montant > 0, vide)
        credit = (-montant).where(montant < 0, vide)
    if structure.get("inverse"):
        debit, credit = credit, debit
    return debit, credit


def table_normalisee(fichier):
    structure = STRUCTURES[fichier["signature"]]
    if fichier["cache"] and fichier["cache"][0] == structure["version"]:
        return fichier["cache"][1]
    df, m = fichier["df"], structure["mapping"]
    t = pd.DataFrame(index=df.index)
    t["Fichier"] = fichier["nom"]
    if m.get("compte"):
        t["Compte"] = identifiant(df[m["compte"]])
    debit, credit = montants(df, structure)
    if debit is not None:
        t["Débit"] = debit
    if credit is not None:
        t["Crédit"] = credit
    if m.get("piece"):
        t["Pièce"] = identifiant(df[m["piece"]])
    for cle, nom in (("libelle", "Libellé"), ("impute", "Imputé par"),
                     ("autorise", "Autorisé par")):
        if m.get(cle):
            t[nom] = df[m[cle]]
    if m.get("date"):
        t["Date de transaction"] = vers_date(df[m["date"]])
    if m.get("date_valeur"):
        t["Date de valeur"] = vers_date(df[m["date_valeur"]])
    t["Devise"] = serie_devise(df[m["devise"]]) if m.get("devise") else "XOF"
    fichier["cache"] = (structure["version"], t)
    return t


def selection(ids):
    return [FICHIERS[i] for i in ids if i in FICHIERS]


# =====================================================================
#  5. Analyse
# =====================================================================

def format_fr(nombre, decimales=0):
    try:
        texte = ("{:,.%df}" % decimales).format(float(nombre))
    except (TypeError, ValueError):
        return str(nombre)
    return texte.replace(",", "\u202f").replace(".", ",")


def valeur_json(valeur):
    if valeur is None:
        return None
    if isinstance(valeur, (pd.Timestamp, datetime, date)):
        return None if pd.isna(valeur) else valeur.strftime("%d/%m/%Y")
    if isinstance(valeur, (bool, np.bool_)):
        return bool(valeur)
    if isinstance(valeur, (int, np.integer)):
        return int(valeur)
    if isinstance(valeur, (float, np.floating)):
        return None if math.isnan(float(valeur)) else float(valeur)
    try:
        if pd.isna(valeur):
            return None
    except (TypeError, ValueError):
        pass
    return str(valeur)


def apercu(serie, nombre=4, montant=False):
    """Quelques valeurs distinctes. Si le début de la colonne est
    monotone (952 sur 3 000 lignes), on cherche plus loin."""
    valeurs = serie.dropna()
    if montant:
        return [format_fr(v) for v in valeurs.head(nombre)]
    vus = []
    for source in (valeurs.head(3000), pd.unique(valeurs)):
        for v in source:
            texte = v.strftime("%d/%m/%Y") if isinstance(v, pd.Timestamp) else str(v).strip().replace("\n", " ")
            court = texte[:42]
            if court and court not in vus:
                vus.append(court)
            if len(vus) >= nombre:
                return vus
    return vus


def analyser(ids):
    fichiers = selection(ids)
    if not fichiers:
        return {"structures": [], "controles": [], "devises": [], "colonnes": [],
                "lignes": 0, "manquants": []}

    groupes = {}
    for f in fichiers:
        groupes.setdefault(f["signature"], []).append(f)

    structures, manquants_globaux = [], []
    for sig, groupe in groupes.items():
        s = STRUCTURES[sig]
        m = s["mapping"]
        table = pd.concat([table_normalisee(f) for f in groupe], ignore_index=True)
        total = int(len(table))
        derive = s["mode"] == "montant"

        def champ(cle, nom_table, requis=True):
            # En montant unique, Débit et Crédit se règlent tous deux par la
            # colonne du montant.
            choix = "montant" if derive and cle in ("debit", "credit") else cle
            present = nom_table in table.columns
            if cle == "devise":
                statut = "present" if m.get("devise") else "neutre"
            else:
                statut = "present" if present else ("absent" if requis else "neutre")
            entree = {"cle": cle, "libelle": LIBELLES[cle], "choix": choix,
                      "colonne": m.get(choix), "statut": statut, "total": total,
                      "manquantes": None}
            if present and cle != "devise":
                entree["manquantes"] = int(table[nom_table].isna().sum())
            return entree

        champs = [champ("compte", "Compte"), champ("debit", "Débit"), champ("credit", "Crédit")]
        # Une écriture est au débit ou au crédit : seule la ligne vide des
        # deux côtés est réellement manquante.
        if "Débit" in table.columns and "Crédit" in table.columns:
            vides = int((table["Débit"].isna() & table["Crédit"].isna()).sum())
            for c in champs[1:]:
                c["manquantes"] = vides
        for cle, nom in (("piece", "Pièce"), ("libelle", "Libellé"),
                         ("impute", "Imputé par"), ("autorise", "Autorisé par")):
            champs.append(champ(cle, nom))
        champs.append(champ("devise", "Devise", requis=False))

        manquants = [c["libelle"] for c in champs if c["statut"] == "absent"]
        manquants_globaux += [x for x in manquants if x not in manquants_globaux]

        # Aperçu des colonnes, montré seulement dans la fenêtre de choix,
        # tiré de tous les fichiers sélectionnés de la structure.
        colonnes = []
        for col in groupe[0]["df"].columns:
            vus = []
            for f in groupe:
                vus += [v for v in f["apercus"].get(str(col), []) if v not in vus]
                if len(vus) >= 4:
                    break
            colonnes.append({"nom": str(col), "exemples": vus[:4]})

        structures.append({
            "signature": sig, "mode": s["mode"], "inverse": bool(s.get("inverse")),
            "montant": m.get("montant"), "lettres": bool(derive and m.get("sens")),
            "fichiers": [f["nom"] for f in groupe], "lignes": total,
            "champs": champs, "manquants": manquants, "colonnes": colonnes,
        })

    table = pd.concat([table_normalisee(f) for f in fichiers], ignore_index=True)
    controles = [{"intitule": "Volumétrie", "valeur": "%s lignes" % format_fr(len(table)),
                  "note": "%d fichier%s" % (len(fichiers), "s" if len(fichiers) > 1 else ""),
                  "niveau": "info"}]
    date = "Date de transaction"
    if date in table.columns and table[date].notna().any():
        controles.append({"intitule": "Période couverte", "niveau": "info", "note": None,
                          "valeur": "%s – %s" % (table[date].min().strftime("%d/%m/%Y"),
                                                 table[date].max().strftime("%d/%m/%Y"))})
    if "Débit" in table.columns and "Crédit" in table.columns:
        ecart = float(table["Débit"].fillna(0).sum() - table["Crédit"].fillna(0).sum())
        controles.append({"intitule": "Équilibre débit / crédit", "note": None,
                          "valeur": "écart de %s" % format_fr(ecart),
                          "niveau": "ok" if abs(ecart) < 0.005 else "alerte"})
    coupes = [f for f in fichiers if f.get("coupure")]
    if coupes:
        premiere = coupes[0]["coupure"]
        controles.append({
            "intitule": "Dernière ligne incomplète", "niveau": "alerte",
            "valeur": ("%d champs sur %d" % (premiere["champs"], premiere["attendus"])
                       if len(coupes) == 1 else "%d fichiers" % len(coupes)),
            "note": ("Ligne conservée, le fichier s'arrête au milieu de celle-ci."
                     if len(fichiers) == 1 else
                     "Lignes conservées : " + ", ".join(f["nom"] for f in coupes))})
    comptes_devises = table["Devise"].value_counts()
    controles.append({"intitule": "Devises", "niveau": "info", "note": None,
                      "valeur": ", ".join(comptes_devises.index[:8])
                      + (" (+%d)" % (len(comptes_devises) - 8) if len(comptes_devises) > 8 else "")})

    return {
        "structures": structures,
        "controles": controles,
        "manquants": manquants_globaux,
        "lignes": int(len(table)),
        "devises": [{"code": str(c), "lignes": int(n)} for c, n in comptes_devises.items()],
        "colonnes": [c for c in COLONNES_TABLE if c in table.columns],
    }


# =====================================================================
#  6. Extraction
# =====================================================================

def masque_booleen(serie):
    return serie.fillna(False).astype(bool).to_numpy()


def lire_montant(texte):
    """Montant saisi à la française : « 30 000 », « 30000,50 », « 1.234,5 »."""
    t = re.sub(r"[\s\u00a0\u202f]", "", str(texte or ""))
    if not t:
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return abs(float(t))
    except ValueError:
        return None


def appliquer_filtres(table, f):
    masque = np.ones(len(table), dtype=bool)
    criteres = []

    def absente(etiquette):
        nonlocal masque
        masque &= False
        criteres.append("%s : colonne absente" % etiquette)

    compte = (f.get("compte") or "").strip()
    if compte:
        if "Compte" in table.columns:
            masque &= masque_booleen(table["Compte"].str.startswith(compte))
            criteres.append("compte commençant par %s" % compte)
        else:
            absente("compte")

    brut = (f.get("montant") or "").strip()
    if brut:
        valeur = lire_montant(brut)
        if valeur is None:
            criteres.append("montant illisible : %s" % brut)
            masque &= False
        elif "Débit" in table.columns or "Crédit" in table.columns:
            # Le montant saisi est cherché au débit comme au crédit.
            trouve = np.zeros(len(table), dtype=bool)
            for col in ("Débit", "Crédit"):
                if col in table.columns:
                    trouve |= (table[col].round(2) == round(valeur, 2)).fillna(False).to_numpy()
            masque &= trouve
            criteres.append("montant %s" % format_fr(valeur, 0 if valeur.is_integer() else 2))
        else:
            absente("montant")

    piece = (f.get("piece") or "").strip()
    if piece:
        if "Pièce" in table.columns:
            masque &= masque_booleen(table["Pièce"] == piece)
            criteres.append("pièce %s" % piece)
        else:
            absente("pièce")

    for cle, colonne in (("libelle", "Libellé"), ("impute", "Imputé par"),
                         ("autorise", "Autorisé par")):
        valeur = (f.get(cle) or "").strip()
        if not valeur:
            continue
        etiquette = colonne.lower()
        if colonne not in table.columns:
            absente(etiquette)
            continue
        contient = lambda s, v=valeur: s.astype("string").str.contains(re.escape(v), case=False, regex=True)
        masque &= masque_booleen(par_valeurs_uniques(table[colonne], contient))
        criteres.append("%s contenant « %s »" % (etiquette, valeur))

    for cle, colonne in (("date_transaction", "Date de transaction"),
                         ("date_valeur", "Date de valeur")):
        brut = (f.get(cle) or "").strip()
        if not brut:
            continue
        if colonne not in table.columns:
            absente(colonne.lower())
            continue
        jour = pd.to_datetime(brut, errors="coerce")
        if pd.notna(jour):
            masque &= masque_booleen(table[colonne].dt.normalize() == jour.normalize())
            criteres.append("%s %s" % (colonne.lower(), jour.strftime("%d/%m/%Y")))

    devises = [d for d in (f.get("devises") or []) if d]
    if devises:
        masque &= table["Devise"].isin(devises).to_numpy()
        criteres.append("devise %s" % ", ".join(devises))

    return table[masque].reset_index(drop=True), criteres


def cle_resultat(ids, filtres):
    versions = [(i, STRUCTURES[FICHIERS[i]["signature"]]["version"]) for i in ids if i in FICHIERS]
    return repr((versions, sorted((k, repr(v)) for k, v in (filtres or {}).items())))


def resultat(ids, filtres, memoriser=True):
    cle = cle_resultat(ids, filtres)
    if RESULTAT["cle"] == cle:
        return RESULTAT["table"], RESULTAT["criteres"]
    fichiers = selection(ids)
    table = pd.concat([table_normalisee(f) for f in fichiers], ignore_index=True)
    table = table[[c for c in COLONNES_TABLE if c in table.columns]]
    table, criteres = appliquer_filtres(table, filtres or {})
    if memoriser:
        RESULTAT.update({"cle": cle, "table": table, "criteres": criteres})
    return table, criteres


def lignes_json(table, debut, fin):
    morceau = table.iloc[debut:fin]
    return [[valeur_json(v) for v in ligne] for ligne in morceau.itertuples(index=False, name=None)]


# =====================================================================
#  7. Routes
# =====================================================================

@app.route("/")
def accueil():
    return Response(PAGE, mimetype="text/html; charset=utf-8")


@app.route("/api/fichiers", methods=["POST"])
def ajouter_fichier():
    nom = (request.args.get("nom") or "fichier").strip()
    ext = os.path.splitext(nom)[1].lower()
    if ext not in EXTENSIONS:
        return jsonify({"erreur": "« %s » : format non pris en charge. Formats acceptés : %s."
                                  % (nom, ", ".join(EXTENSIONS))}), 400

    fid = uuid.uuid4().hex[:12]
    source = os.path.join(DOSSIER_TRAVAIL, fid + ext)
    # Écriture par blocs : un fichier volumineux ne passe jamais en entier
    # par la mémoire pendant l'envoi.
    taille = 0
    with open(source, "wb") as sortie:
        while True:
            bloc = request.stream.read(8 * 1024 * 1024)
            if not bloc:
                break
            sortie.write(bloc)
            taille += len(bloc)
    if taille == 0:
        os.remove(source)
        return jsonify({"erreur": "« %s » est vide." % nom}), 400

    affiche = nom
    existants = {f["nom"] for f in FICHIERS.values()}
    n = 2
    while affiche in existants:
        racine, e = os.path.splitext(nom)
        affiche = "%s (%d)%s" % (racine, n, e)
        n += 1

    fichier = {"id": fid, "nom": affiche, "source": source, "taille": taille,
               "feuilles": [], "feuille": None}
    try:
        with VERROU:
            if ext in (".xlsx", ".xlsm", ".xls"):
                fichier["csv"] = os.path.join(DOSSIER_TRAVAIL, fid + ".csv")
                fichier["feuilles"], fichier["feuille"] = convertir_excel(source, fichier["csv"])
            else:
                fichier["csv"] = source
            charger(fichier)
            FICHIERS[fid] = fichier
    except ValueError as err:
        return jsonify({"erreur": "« %s » : %s" % (nom, err)}), 400
    except Exception as err:
        return jsonify({"erreur": "« %s » illisible : %s" % (nom, err)}), 400
    return jsonify({"fichier": infos_fichier(fichier)})


@app.route("/api/fichiers", methods=["GET"])
def lister_fichiers():
    """Fichiers déjà chargés : la page les retrouve après un rechargement."""
    with VERROU:
        return jsonify({"fichiers": [infos_fichier(f) for f in FICHIERS.values()]})


@app.route("/api/fichiers/<fid>", methods=["DELETE"])
def retirer_fichier(fid):
    with VERROU:
        fichier = FICHIERS.pop(fid, None)
        if fichier:
            for chemin in {fichier["source"], fichier.get("csv")}:
                if chemin and os.path.exists(chemin):
                    os.remove(chemin)
            RESULTAT.update({"cle": None, "table": None})
    return jsonify({"ok": True})


@app.route("/api/feuille", methods=["POST"])
def changer_feuille():
    d = request.get_json(silent=True) or {}
    fichier = FICHIERS.get(d.get("id"))
    if not fichier or not fichier["feuilles"]:
        return jsonify({"erreur": "Fichier introuvable."}), 400
    try:
        with VERROU:
            fichier["feuilles"], fichier["feuille"] = convertir_excel(
                fichier["source"], fichier["csv"], d.get("feuille"))
            charger(fichier)
    except Exception as err:
        return jsonify({"erreur": "Lecture de la feuille impossible : %s" % err}), 400
    return jsonify({"fichier": infos_fichier(fichier)})


@app.route("/api/analyse", methods=["POST"])
def route_analyse():
    d = request.get_json(silent=True) or {}
    with VERROU:
        return jsonify(analyser(d.get("fichiers", [])))


@app.route("/api/correspondance", methods=["POST"])
def route_correspondance():
    d = request.get_json(silent=True) or {}
    structure = STRUCTURES.get(d.get("signature"))
    if structure is None:
        return jsonify({"erreur": "Structure inconnue."}), 400
    with VERROU:
        if d.get("mode") in ("dc", "montant"):
            structure["mode"] = d["mode"]
        if isinstance(d.get("inverse"), bool):
            structure["inverse"] = d["inverse"]
        if d.get("cle") in CLES:
            colonne = d.get("colonne")
            if colonne is not None and colonne not in structure["colonnes"]:
                return jsonify({"erreur": "Colonne inconnue : %s" % colonne}), 400
            structure["mapping"][d["cle"]] = colonne
        structure["version"] += 1
        return jsonify(analyser(d.get("fichiers", [])))


@app.route("/api/extraire", methods=["POST"])
def route_extraire():
    d = request.get_json(silent=True) or {}
    ids = d.get("fichiers", [])
    if not selection(ids):
        return jsonify({"erreur": "Aucun fichier sélectionné."}), 400
    with VERROU:
        table, criteres = resultat(ids, d.get("filtres"))
        return jsonify({"total": int(len(table)), "colonnes": list(table.columns),
                        "criteres": criteres, "lignes": lignes_json(table, 0, 200)})


@app.route("/api/lignes", methods=["POST"])
def route_lignes():
    d = request.get_json(silent=True) or {}
    with VERROU:
        table = RESULTAT["table"]
        if table is None:
            return jsonify({"erreur": "Relancez l'extraction."}), 400
        debut = max(0, int(d.get("debut", 0)))
        fin = min(len(table), int(d.get("fin", debut + 200)))
        return jsonify({"debut": debut, "lignes": lignes_json(table, debut, fin)})


def valeur_excel(valeur):
    if valeur is None:
        return None
    if isinstance(valeur, pd.Timestamp):
        return None if pd.isna(valeur) else valeur.to_pydatetime()
    try:
        if pd.isna(valeur):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valeur, np.generic):
        return valeur.item()
    return valeur


@app.route("/api/exporter", methods=["POST"])
def route_exporter():
    d = request.get_json(silent=True) or {}
    ids = d.get("fichiers", [])
    if not selection(ids):
        return jsonify({"erreur": "Aucun fichier sélectionné."}), 400
    with VERROU:
        table, _ = resultat(ids, d.get("filtres"), memoriser=False)
        colonnes = [c for c in (d.get("colonnes") or []) if c in table.columns] or list(table.columns)
        table = table[colonnes].copy()
    # Montants entiers écrits sans décimale : 30000 et non 30000,0.
    for col in ("Débit", "Crédit"):
        if col in table.columns:
            valeurs = table[col].dropna()
            if len(valeurs) and bool((valeurs % 1 == 0).all()):
                table[col] = table[col].astype("Int64")

    if d.get("format") == "csv":
        tampon = io.BytesIO()
        tampon.write(table.to_csv(index=False, sep=";", decimal=",", date_format="%d/%m/%Y")
                     .encode("utf-8-sig"))
        tampon.seek(0)
        return send_file(tampon, mimetype="text/csv", as_attachment=True,
                         download_name="extraction.csv")

    if len(table) > 1048575:
        return jsonify({"erreur": "%s lignes : Excel ne peut pas en contenir plus de 1 048 575. "
                                  "Exportez en CSV." % format_fr(len(table))}), 400
    from openpyxl import Workbook
    classeur = Workbook(write_only=True)
    feuille = classeur.create_sheet("Extraction")
    feuille.append(colonnes)
    for ligne in table.itertuples(index=False, name=None):
        feuille.append([valeur_excel(v) for v in ligne])
    tampon = io.BytesIO()
    classeur.save(tampon)
    tampon.seek(0)
    return send_file(tampon, as_attachment=True, download_name="extraction.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# =====================================================================
#  8. Interface
# =====================================================================

PAGE = r"""<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Grand livre</title>
<script>
  // Thème appliqué avant l'affichage, pour éviter un flash clair en mode sombre.
  try {
    var t = localStorage.getItem('theme');
    if (!t && window.matchMedia('(prefers-color-scheme: dark)').matches) t = 'sombre';
    if (t === 'sombre') document.documentElement.dataset.theme = 'sombre';
  } catch (e) {}
</script>
<style>
  :root{
    color-scheme:light;
    --page:#F5F5F7; --surface:#FFFFFF; --champ:#FFFFFF; --texte:#1D1D1F; --second:#6E6E73;
    --tiers:#8E8E93; --filet:#D2D2D7; --filet-doux:#E8E8ED; --remplissage:#F2F2F5;
    --accent:#1D1D1F; --sur-accent:#FFFFFF; --accent-survol:#3A3A3C;
    --doux:#E8E8ED; --doux-survol:#DCDCE1; --bord:#C7C7CC; --bord-survol:#AEAEB2;
    --vert:#248A3D; --rouge:#D70015; --rouge-fond:#FFF2F2; --rouge-filet:#F5C9C9;
    --bandeau:#000000; --bandeau-texte:#D1D1D6;
    --entete:#000000; --entete-texte:#FFFFFF; --rayure:#FAFAFB; --survol-depot:#FAFAFA;
    --pop:rgba(255,255,255,.97); --ombre-pop:0 12px 40px rgba(0,0,0,.16), 0 0 0 .5px rgba(0,0,0,.14);
    --ombre-segment:0 1px 3px rgba(0,0,0,.12), 0 0 0 .5px rgba(0,0,0,.04);
    --piste:#E8E8ED; --piste-choix:#FFFFFF;
    --mono:ui-monospace,"SF Mono","Cascadia Mono",Consolas,monospace;
  }
  :root[data-theme="sombre"]{
    color-scheme:dark;
    --page:#000000; --surface:#1C1C1E; --champ:#2C2C2E; --texte:#F5F5F7; --second:#98989D;
    --tiers:#8E8E93; --filet:#38383A; --filet-doux:#2C2C2E; --remplissage:#2C2C2E;
    --accent:#F5F5F7; --sur-accent:#000000; --accent-survol:#D1D1D6;
    --doux:#2C2C2E; --doux-survol:#3A3A3C; --bord:#48484A; --bord-survol:#636366;
    --vert:#30D158; --rouge:#FF453A; --rouge-fond:rgba(255,69,58,.12); --rouge-filet:rgba(255,69,58,.32);
    --bandeau:#1C1C1E; --bandeau-texte:#AEAEB2;
    --entete:#2C2C2E; --entete-texte:#F5F5F7; --rayure:#232325; --survol-depot:#232325;
    --pop:rgba(44,44,46,.97); --ombre-pop:0 12px 40px rgba(0,0,0,.55), 0 0 0 .5px rgba(255,255,255,.1);
    --ombre-segment:0 1px 3px rgba(0,0,0,.5);
    --piste:rgba(118,118,128,.24); --piste-choix:#636366;
  }
  *{box-sizing:border-box}
  html{-webkit-text-size-adjust:100%}
  body{margin:0; background:var(--page); color:var(--texte);
       font-family:-apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",system-ui,Roboto,"Helvetica Neue",Arial,sans-serif;
       font-size:15px; line-height:1.45; letter-spacing:-.01em; -webkit-font-smoothing:antialiased;
       transition:background .25s, color .25s}
  button,input,select{font:inherit; color:inherit; letter-spacing:inherit}
  button{cursor:pointer}
  :focus-visible{outline:3px solid color-mix(in srgb, var(--texte) 35%, transparent); outline-offset:2px; border-radius:6px}
  .chiffres{font-variant-numeric:tabular-nums}
  .cache{display:none !important}
  .i-ok{color:var(--vert)} .i-ko{color:var(--rouge)} .i-neutre{color:var(--tiers)}
  svg{display:block}

  header{position:sticky; top:0; z-index:20; background:var(--bandeau); color:var(--bandeau-texte);
         padding:9px 20px 9px 28px; font-size:13px; display:flex; gap:20px; align-items:center;
         border-bottom:.5px solid var(--filet)}
  .theme{margin-left:auto; border:0; background:none; color:#fff; width:30px; height:30px;
         border-radius:50%; display:grid; place-items:center; opacity:.85}
  .theme:hover{opacity:1; background:rgba(255,255,255,.12)}
  main{max-width:1200px; margin:0 auto; padding:28px 24px 96px}

  section{background:var(--surface); border-radius:18px; margin-bottom:22px; overflow:hidden}
  .tete{padding:20px 24px 6px; display:flex; align-items:center; gap:14px; flex-wrap:wrap}
  .tete h2{margin:0; font-size:21px; font-weight:600; letter-spacing:-.022em}
  .tete p{margin:0; color:var(--second); font-size:13px}
  .corps{padding:12px 24px 24px}

  .depot{margin:24px; border:1.5px dashed var(--bord); border-radius:14px; padding:40px 20px;
         text-align:center; transition:border-color .2s, background .2s}
  .depot.survol{border-color:var(--texte); background:var(--survol-depot)}
  .depot h2{margin:0 0 6px; font-size:21px; font-weight:600; letter-spacing:-.022em}
  .depot p{margin:0 0 20px; color:var(--second); font-size:14px}

  .bouton{border:0; border-radius:980px; padding:8px 18px; font-size:14px; font-weight:500;
          background:var(--accent); color:var(--sur-accent); transition:background .15s, opacity .15s}
  .bouton:hover{background:var(--accent-survol)}
  .bouton.doux{background:var(--doux); color:var(--texte)}
  .bouton.doux:hover{background:var(--doux-survol)}
  .bouton:disabled{opacity:.4; cursor:default}
  .lien{border:0; background:none; padding:0; color:var(--texte); font-size:13px;
        font-weight:500; text-decoration:underline; text-underline-offset:3px}

  .envois{padding:0 24px 20px}
  .envoi{display:grid; grid-template-columns:1fr auto; gap:6px 16px; padding:10px 0; border-top:.5px solid var(--filet)}
  .envoi .nom{font-size:14px; font-weight:500; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
  .envoi .statut{font-size:13px; color:var(--second); text-align:right}
  .envoi .statut.ko{color:var(--rouge)}
  .jauge{grid-column:1 / -1; height:4px; border-radius:2px; background:var(--filet-doux); overflow:hidden; position:relative}
  .jauge i{position:absolute; inset:0 auto 0 0; width:0; background:var(--accent); border-radius:2px; transition:width .25s ease}
  .jauge.continue i{width:32%; animation:glisse 1.1s ease-in-out infinite}
  .jauge.finie i{width:100%}
  @keyframes glisse{0%{left:-32%} 100%{left:100%}}

  .liste-fichiers{list-style:none; margin:0; padding:0}
  .ligne-fichier{display:grid; grid-template-columns:auto 1fr auto auto auto; gap:14px; align-items:center;
                 padding:11px 0; border-top:.5px solid var(--filet)}
  .ligne-fichier:first-child{border-top:0}
  .rond{width:22px; height:22px; border-radius:50%; border:1.5px solid var(--bord); background:var(--champ);
        display:grid; place-items:center; padding:0; color:var(--sur-accent)}
  .rond svg{opacity:0}
  .rond[aria-checked="true"]{background:var(--accent); border-color:var(--accent)}
  .rond[aria-checked="true"] svg{opacity:1}
  .ligne-fichier .nom{font-weight:500; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
  .ligne-fichier .meta{color:var(--second); font-size:13px}
  .ligne-fichier select{border:.5px solid var(--filet); border-radius:8px; padding:3px 8px; font-size:13px; background:var(--champ)}
  .retirer{border:0; background:none; color:var(--tiers); padding:4px; border-radius:6px; display:grid}
  .retirer:hover{color:var(--texte); background:var(--remplissage)}

  .onglets{display:flex; gap:6px; flex-wrap:wrap; margin-left:auto}
  .rappel{display:flex; gap:10px; align-items:center; background:var(--rouge-fond); border:.5px solid var(--rouge-filet);
          border-radius:12px; padding:11px 14px; margin:6px 0 10px; color:var(--rouge); font-weight:600; font-size:14px}
  .champ{display:grid; grid-template-columns:24px 128px 1fr auto; gap:12px; align-items:center;
         padding:11px 0; border-top:.5px solid var(--filet-doux)}
  .groupe > .champ:first-child{border-top:0}
  .champ .nom-champ{font-weight:500}
  .champ.absent{background:var(--rouge-fond); border-radius:12px; margin:6px -12px; padding:11px 12px; border-top-color:transparent}
  .champ.absent .nom-champ{color:var(--rouge)}
  .champ.absent + .champ{border-top-color:transparent}
  .champ .manque{font-size:13px; color:var(--second); white-space:nowrap}
  .choix-zone{display:flex; align-items:center; gap:10px}
  .choix{display:inline-flex; align-items:center; justify-content:space-between; gap:10px;
         min-width:190px; max-width:280px; background:var(--champ); border:.5px solid var(--filet);
         border-radius:8px; padding:5px 8px 5px 10px; box-shadow:0 1px 2px rgba(0,0,0,.05); text-align:left}
  .choix:hover{border-color:var(--bord-survol)}
  .choix span{font-family:var(--mono); font-size:13px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
  .choix span.aucune{font-family:inherit; color:var(--tiers)}
  .choix svg{flex:none; color:var(--second)}
  .signe{width:30px; height:30px; border-radius:50%; border:.5px solid var(--filet); background:var(--champ);
         font-size:17px; font-weight:600; line-height:1; display:grid; place-items:center; padding:0;
         box-shadow:0 1px 2px rgba(0,0,0,.05); font-variant-numeric:tabular-nums}
  .signe:hover{border-color:var(--bord-survol)}
  .bascule{padding:12px 0; border-top:.5px solid var(--filet); border-bottom:.5px solid var(--filet-doux)}
  .groupe + .groupe{border-top:.5px solid var(--filet)}
  .segment{display:inline-flex; background:var(--piste); border-radius:9px; padding:2px}
  .segment button{border:0; background:none; border-radius:7px; padding:5px 14px; font-size:13px; font-weight:500; color:var(--texte)}
  .segment button[aria-pressed="true"]{background:var(--piste-choix); box-shadow:var(--ombre-segment)}

  .voile{position:fixed; inset:0; z-index:40}
  .popover{position:fixed; z-index:41; width:380px; max-width:calc(100vw - 24px); max-height:min(460px, 70vh);
           display:flex; flex-direction:column; background:var(--pop);
           -webkit-backdrop-filter:saturate(180%) blur(20px); backdrop-filter:saturate(180%) blur(20px);
           border-radius:14px; box-shadow:var(--ombre-pop); transform-origin:top center; animation:ouvre .16s ease-out}
  @keyframes ouvre{from{opacity:0; transform:scale(.97) translateY(-4px)} to{opacity:1; transform:none}}
  .popover .titre-pop{padding:14px 16px 8px; font-size:13px; color:var(--second)}
  .popover .titre-pop strong{color:var(--texte); font-weight:600}
  .popover input{margin:0 12px 8px; padding:7px 10px; border:0; border-radius:8px; background:var(--remplissage); font-size:14px}
  .options{overflow:auto; padding:0 6px 8px}
  .option{display:grid; grid-template-columns:18px 1fr; gap:8px; width:100%; text-align:left;
          border:0; background:none; border-radius:9px; padding:8px 10px}
  .option:hover,.option:focus-visible{background:var(--remplissage); outline:none}
  .option .coche{padding-top:2px; visibility:hidden}
  .option[aria-selected="true"] .coche{visibility:visible}
  .option .nom-col{font-family:var(--mono); font-size:13px; font-weight:600}
  .option .aucune{font-weight:500}
  .option .ex{display:flex; gap:4px; flex-wrap:wrap; margin-top:4px}
  .option .ex span{font-size:12px; background:var(--surface); border:.5px solid var(--filet-doux); border-radius:5px;
                   padding:1px 6px; color:var(--second); max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}

  .controles{display:grid; grid-template-columns:repeat(auto-fill,minmax(230px,1fr)); gap:12px}
  .controle{border:.5px solid var(--filet); border-radius:12px; padding:13px 15px}
  .controle .intitule{font-size:13px; color:var(--second)}
  .controle .valeur{font-size:17px; font-weight:500; margin-top:2px}
  .controle .note{font-size:12px; color:var(--second); margin-top:2px}
  .controle.alerte .valeur{color:var(--rouge)}

  .grille{display:grid; grid-template-columns:repeat(auto-fill,minmax(200px,1fr)); gap:14px}
  label{display:block; font-size:13px; color:var(--second); margin-bottom:5px}
  input[type=text],input[type=date],input[type=search]{width:100%; padding:7px 10px; border:.5px solid var(--filet);
                                    border-radius:8px; background:var(--champ)}
  .rangee{margin-top:18px}
  .rangee > .etiquette{font-size:13px; color:var(--second); margin-bottom:7px}
  .puces{display:flex; gap:6px; flex-wrap:wrap}
  .puce{border:.5px solid var(--filet); background:var(--champ); border-radius:980px; padding:5px 12px; font-size:13px; font-weight:500}
  .puce[aria-pressed="true"]{background:var(--accent); border-color:var(--accent); color:var(--sur-accent)}
  .puce .n{font-weight:400; opacity:.65; margin-left:5px}
  .actions{display:flex; gap:10px; margin-top:20px; flex-wrap:wrap; align-items:center}
  .bilan{font-size:13px; color:var(--second)}
  .criteres{margin-top:12px; display:flex; gap:6px; flex-wrap:wrap}
  .critere{font-size:12px; border:.5px solid var(--filet); border-radius:980px; padding:2px 10px; color:var(--second)}
  .erreur{background:var(--rouge-fond); border:.5px solid var(--rouge-filet); color:var(--rouge);
          padding:11px 14px; border-radius:12px; font-size:14px; margin-bottom:18px}

  .cadre{margin-top:12px; height:560px; overflow:auto; border:.5px solid var(--filet); border-radius:12px}
  table{border-collapse:separate; border-spacing:0; table-layout:fixed; font-size:13px}
  th,td{height:36px; padding:0 12px; text-align:left; white-space:nowrap; overflow:hidden;
        text-overflow:ellipsis; box-shadow:inset 0 -.5px 0 var(--filet-doux)}
  thead th{position:sticky; top:0; z-index:1; background:var(--entete); color:var(--entete-texte); font-weight:500; height:38px}
  tbody tr.pair td{background:var(--rayure)}
  th.nombre,td.nombre{text-align:right; font-variant-numeric:tabular-nums}
  tr.espace td{padding:0; box-shadow:none; height:auto}
  tr.attente td{color:var(--tiers)}

  @media (prefers-reduced-motion:reduce){
    .jauge.continue i{animation:none; width:100%; opacity:.35}
    .popover{animation:none} body{transition:none}
  }
  @media (max-width:760px){
    main{padding:18px 12px 72px}
    .champ{grid-template-columns:24px 1fr}
    .champ > .choix-zone,.champ > .manque{grid-column:2}
    .ligne-fichier{grid-template-columns:auto 1fr auto}
    .ligne-fichier .meta{grid-column:2}
  }
</style>
</head>
<body>
<header>
  <span class="etat chiffres" id="etat-general">Aucun fichier chargé</span>
  <button class="theme" id="btn-theme" aria-label="Activer le mode sombre"></button>
</header>

<main>
  <div id="zone-erreur"></div>

  <section>
    <div class="depot" id="depot">
      <h2>Déposez vos grands livres</h2>
      <p>Un ou plusieurs fichiers.</p>
      <button class="bouton" id="btn-parcourir">Choisir des fichiers</button>
      <input type="file" id="input-fichiers" class="cache" multiple accept=".xlsx,.xlsm,.xls,.csv,.txt,.tsv">
    </div>
    <div class="envois" id="envois"></div>
  </section>

  <section id="bloc-fichiers" class="cache">
    <div class="tete">
      <h2>Fichiers</h2>
      <p class="chiffres" id="compte-selection"></p>
      <button class="lien" id="btn-tout" style="margin-left:auto">Tout sélectionner</button>
    </div>
    <div class="corps"><ul class="liste-fichiers" id="liste-fichiers"></ul></div>
  </section>

  <div id="bloc-analyse" class="cache">
    <section>
      <div class="tete">
        <h2>Colonnes attendues</h2>
        <div class="onglets" id="onglets-structures"></div>
      </div>
      <div class="corps" id="correspondance"></div>
    </section>

    <section>
      <div class="tete"><h2>État du fichier</h2></div>
      <div class="corps"><div class="controles" id="liste-controles"></div></div>
    </section>

    <section>
      <div class="tete">
        <h2>Extraction</h2>
        <p>Laissez vide ce que vous ne voulez pas filtrer.</p>
      </div>
      <div class="corps">
        <div class="grille">
          <div><label for="f-compte">Compte</label><input type="text" id="f-compte" placeholder="401100" autocomplete="off"></div>
          <div><label for="f-montant">Montant</label><input type="text" id="f-montant" placeholder="30 000" inputmode="decimal" autocomplete="off"></div>
          <div><label for="f-piece">Pièce</label><input type="text" id="f-piece" placeholder="797" autocomplete="off"></div>
          <div><label for="f-libelle">Libellé</label><input type="text" id="f-libelle" placeholder="provision" autocomplete="off"></div>
          <div><label for="f-impute">Imputé par</label><input type="text" id="f-impute" placeholder="nom de l'utilisateur" autocomplete="off"></div>
          <div><label for="f-autorise">Autorisé par</label><input type="text" id="f-autorise" placeholder="nom du validateur" autocomplete="off"></div>
          <div><label for="f-date-transaction">Date de transaction</label><input type="date" id="f-date-transaction"></div>
          <div><label for="f-date-valeur">Date de valeur</label><input type="date" id="f-date-valeur"></div>
        </div>
        <div class="rangee">
          <div class="etiquette">Devise</div>
          <div class="puces" id="puces-devises"></div>
        </div>
        <div class="rangee">
          <div class="etiquette">Colonnes affichées</div>
          <div class="puces" id="puces-colonnes"></div>
        </div>
        <div class="actions">
          <button class="bouton" id="btn-extraire">Extraire</button>
          <button class="bouton doux" id="btn-effacer">Tout effacer</button>
          <button class="bouton doux" id="btn-excel" disabled>Exporter en Excel</button>
          <button class="bouton doux" id="btn-csv" disabled>Exporter en CSV</button>
          <span class="bilan chiffres" id="bilan"></span>
        </div>
        <div class="criteres" id="criteres"></div>
        <div class="cadre cache" id="cadre">
          <table><colgroup id="cols"></colgroup><thead id="entete"></thead><tbody id="corps"></tbody></table>
        </div>
      </div>
    </section>
  </div>
</main>

<script>
const ICONES = {
  ok: '<svg class="i-ok" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-label="présente"><circle cx="12" cy="12" r="9"/><path d="m8 12.3 2.7 2.7L16 9.6"/></svg>',
  ko: '<svg class="i-ko" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-label="absente"><path d="M10.3 3.6 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.6a2 2 0 0 0-3.4 0Z"/><line x1="12" y1="9" x2="12" y2="13.5"/><line x1="12" y1="17.2" x2="12.01" y2="17.2"/></svg>',
  neutre: '<svg class="i-neutre" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-label="facultative"><circle cx="12" cy="12" r="9"/><path d="M8.5 12h7"/></svg>',
  chevrons: '<svg width="10" height="14" viewBox="0 0 10 14" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="m2 5 3-3 3 3M2 9l3 3 3-3"/></svg>',
  coche: '<svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m2.5 7.3 3 3L11.5 4"/></svg>',
  cocheRond: '<svg width="12" height="12" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m2.5 7.3 3 3L11.5 4"/></svg>',
  croix: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="m4.5 4.5 7 7m0-7-7 7"/></svg>',
  lune: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M20.5 14.2A8.5 8.5 0 0 1 9.8 3.5a8.5 8.5 0 1 0 10.7 10.7Z"/></svg>',
  soleil: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="4.2"/><path d="M12 2.5v2.2M12 19.3v2.2M4.6 4.6l1.6 1.6M17.8 17.8l1.6 1.6M2.5 12h2.2M19.3 12h2.2M4.6 19.4l1.6-1.6M17.8 6.2l1.6-1.6"/></svg>'
};
const NUMERIQUES = new Set(['Débit', 'Crédit']);
const LARGEURS = {'Fichier':100, 'Compte':100, 'Débit':120, 'Crédit':120, 'Pièce':100,
                  'Imputé par':120, 'Autorisé par':120, 'Date de transaction':160, 'Date de valeur':128, 'Devise':72};
const NOMS = {compte:'Compte', montant:'Montant', debit:'Débit', credit:'Crédit', piece:'Pièce', libelle:'Libellé',
              impute:'Imputé par', autorise:'Autorisé par', devise:'Devise'};
const HAUTEUR = 36, BLOC = 200;
// Hauteur de défilement maximale, sous les limites des navigateurs.
const HAUTEUR_MAX = 12e6;

const etat = {fichiers: [], selection: new Set(), analyse: null, structure: null,
              colonnesChoisies: null, devises: new Set(), extraction: null};
const $ = id => document.getElementById(id);

function echapper(t){
  return String(t).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function nombreFr(v, dec){
  return new Intl.NumberFormat('fr-FR', {minimumFractionDigits: dec||0, maximumFractionDigits: dec||0}).format(v);
}
function tailleFr(o){
  if (o < 1024) return o + ' o';
  const u = ['Ko','Mo','Go','To']; let i = -1;
  do { o /= 1024; i++; } while (o >= 1024 && i < u.length - 1);
  return nombreFr(o, o < 10 ? 1 : 0) + ' ' + u[i];
}
function erreur(message){
  $('zone-erreur').innerHTML = message ? '<div class="erreur">' + echapper(message) + '</div>' : '';
  if (message) window.scrollTo({top: 0, behavior: 'smooth'});
}
async function api(route, corps){
  const r = await fetch(route, {method: 'POST', headers: {'Content-Type': 'application/json'},
                                body: JSON.stringify(corps || {})});
  const d = await r.json().catch(() => ({erreur: 'Réponse illisible du serveur.'}));
  if (!r.ok) throw new Error(d.erreur || 'Erreur du serveur.');
  return d;
}
const ids = () => etat.fichiers.filter(f => etat.selection.has(f.id)).map(f => f.id);

/* ---------------- thème ---------------- */
function appliquerTheme(sombre){
  if (sombre) document.documentElement.dataset.theme = 'sombre';
  else delete document.documentElement.dataset.theme;
  $('btn-theme').innerHTML = sombre ? ICONES.soleil : ICONES.lune;
  $('btn-theme').setAttribute('aria-label', sombre ? 'Activer le mode clair' : 'Activer le mode sombre');
}
$('btn-theme').addEventListener('click', () => {
  const sombre = document.documentElement.dataset.theme !== 'sombre';
  try { localStorage.setItem('theme', sombre ? 'sombre' : 'clair'); } catch (e) {}
  appliquerTheme(sombre);
});
appliquerTheme(document.documentElement.dataset.theme === 'sombre');

/* ---------------- dépôt et envoi ---------------- */
const depot = $('depot');
$('btn-parcourir').addEventListener('click', () => $('input-fichiers').click());
$('input-fichiers').addEventListener('change', e => { deposer([...e.target.files]); e.target.value = ''; });
['dragenter','dragover'].forEach(t => depot.addEventListener(t, e => { e.preventDefault(); depot.classList.add('survol'); }));
['dragleave','drop'].forEach(t => depot.addEventListener(t, e => { e.preventDefault(); depot.classList.remove('survol'); }));
depot.addEventListener('drop', e => deposer([...e.dataTransfer.files]));

let fileEnvoi = Promise.resolve();
function deposer(liste){
  erreur('');
  liste.forEach(fichier => {
    const ligne = document.createElement('div');
    ligne.className = 'envoi';
    ligne.innerHTML = '<div class="nom">' + echapper(fichier.name) + '</div>'
      + '<div class="statut chiffres">En attente</div><div class="jauge"><i></i></div>';
    $('envois').appendChild(ligne);
    fileEnvoi = fileEnvoi.then(() => envoyer(fichier, ligne));
  });
  fileEnvoi = fileEnvoi.then(() => { if (etat.fichiers.length) analyser(); });
}

function envoyer(fichier, ligne){
  const statut = ligne.querySelector('.statut'), jauge = ligne.querySelector('.jauge'), barre = jauge.querySelector('i');
  return new Promise(resolve => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/fichiers?nom=' + encodeURIComponent(fichier.name));
    xhr.upload.onprogress = e => {
      if (!e.lengthComputable) return;
      const p = e.loaded / e.total;
      barre.style.width = (p * 100).toFixed(1) + '%';
      statut.textContent = 'Envoi ' + Math.floor(p * 100) + ' %, ' + tailleFr(e.total);
    };
    xhr.upload.onload = () => { jauge.classList.add('continue'); statut.textContent = 'Conversion et lecture…'; };
    xhr.onload = () => {
      jauge.classList.remove('continue');
      let d = {};
      try { d = JSON.parse(xhr.responseText); } catch (e) {}
      if (xhr.status === 200 && d.fichier){
        jauge.classList.add('finie');
        statut.textContent = nombreFr(d.fichier.lignes) + ' lignes';
        etat.fichiers.push(d.fichier);
        etat.selection.add(d.fichier.id);
        rendreFichiers();
        setTimeout(() => ligne.remove(), 1600);
      } else {
        statut.classList.add('ko');
        statut.textContent = d.erreur || 'Envoi impossible';
      }
      resolve();
    };
    xhr.onerror = () => { statut.classList.add('ko'); statut.textContent = 'Le serveur ne répond pas'; resolve(); };
    xhr.send(fichier);
  });
}

/* ---------------- fichiers ---------------- */
function rendreFichiers(){
  const n = etat.fichiers.length, s = ids().length;
  $('bloc-fichiers').classList.toggle('cache', n === 0);
  $('compte-selection').textContent = s + ' sélectionné' + (s > 1 ? 's' : '') + ' sur ' + n;
  $('btn-tout').textContent = s === n ? 'Tout désélectionner' : 'Tout sélectionner';
  $('liste-fichiers').innerHTML = etat.fichiers.map(f => {
    const choisi = etat.selection.has(f.id);
    const feuilles = f.feuilles.length > 1
      ? '<select data-feuille="' + f.id + '" aria-label="Feuille">' + f.feuilles.map(x =>
          '<option' + (x === f.feuille ? ' selected' : '') + '>' + echapper(x) + '</option>').join('') + '</select>'
      : '<span></span>';
    return '<li class="ligne-fichier">'
      + '<button class="rond" role="checkbox" aria-checked="' + choisi + '" data-choix="' + f.id + '" aria-label="Sélectionner ' + echapper(f.nom) + '">' + ICONES.cocheRond + '</button>'
      + '<span class="nom" title="' + echapper(f.nom) + '">' + echapper(f.nom) + '</span>'
      + '<span class="meta chiffres">' + nombreFr(f.lignes) + ' lignes, ' + tailleFr(f.taille) + '</span>'
      + feuilles
      + '<button class="retirer" data-retirer="' + f.id + '" aria-label="Retirer ' + echapper(f.nom) + '">' + ICONES.croix + '</button></li>';
  }).join('');
  majEtatGeneral();
}
function majEtatGeneral(){
  const s = ids().length;
  if (!etat.fichiers.length){ $('etat-general').textContent = 'Aucun fichier chargé'; return; }
  const lignes = etat.fichiers.filter(f => etat.selection.has(f.id)).reduce((a, f) => a + f.lignes, 0);
  $('etat-general').textContent = s + ' fichier' + (s > 1 ? 's' : '') + ' sélectionné' + (s > 1 ? 's' : '')
    + ' sur ' + etat.fichiers.length + ', ' + nombreFr(lignes) + ' lignes';
}
$('liste-fichiers').addEventListener('click', async e => {
  const choix = e.target.closest('[data-choix]'), retrait = e.target.closest('[data-retirer]');
  if (choix){
    const id = choix.dataset.choix;
    etat.selection.has(id) ? etat.selection.delete(id) : etat.selection.add(id);
    rendreFichiers(); analyser();
  } else if (retrait){
    const id = retrait.dataset.retirer;
    await fetch('/api/fichiers/' + id, {method: 'DELETE'});
    etat.fichiers = etat.fichiers.filter(f => f.id !== id);
    etat.selection.delete(id);
    rendreFichiers(); analyser();
  }
});
$('liste-fichiers').addEventListener('change', async e => {
  const sel = e.target.closest('[data-feuille]');
  if (!sel) return;
  try {
    const d = await api('/api/feuille', {id: sel.dataset.feuille, feuille: sel.value});
    etat.fichiers = etat.fichiers.map(f => f.id === d.fichier.id ? d.fichier : f);
    rendreFichiers(); analyser();
  } catch (err) { erreur(err.message); }
});
$('btn-tout').addEventListener('click', () => {
  if (ids().length === etat.fichiers.length) etat.selection.clear();
  else etat.fichiers.forEach(f => etat.selection.add(f.id));
  rendreFichiers(); analyser();
});

/* ---------------- analyse ---------------- */
async function analyser(){
  viderExtraction();
  if (!ids().length){ $('bloc-analyse').classList.add('cache'); return; }
  try { rendreAnalyse(await api('/api/analyse', {fichiers: ids()})); }
  catch (err) { erreur(err.message); }
}
async function corriger(charge){
  try {
    rendreAnalyse(await api('/api/correspondance', Object.assign({fichiers: ids(), signature: etat.structure}, charge)));
    viderExtraction();
  } catch (err) { erreur(err.message); }
}
function structureActive(){
  const a = etat.analyse;
  return a.structures.find(s => s.signature === etat.structure) || a.structures[0];
}

function rendreAnalyse(a){
  etat.analyse = a;
  $('bloc-analyse').classList.remove('cache');
  if (!a.structures.some(s => s.signature === etat.structure)) etat.structure = a.structures[0].signature;
  $('onglets-structures').innerHTML = a.structures.length > 1
    ? '<div class="segment">' + a.structures.map((s, i) =>
        '<button data-structure="' + s.signature + '" aria-pressed="' + (s.signature === etat.structure) + '" title="' + echapper(s.fichiers.join(', ')) + '">'
        + 'Structure ' + (i + 1) + ' (' + s.fichiers.length + ')</button>').join('') + '</div>'
    : '';
  rendreCorrespondance();

  $('liste-controles').innerHTML = a.controles.map(c =>
    '<div class="controle ' + c.niveau + '"><div class="intitule">' + echapper(c.intitule) + '</div>'
    + '<div class="valeur chiffres">' + echapper(c.valeur) + '</div>'
    + (c.note ? '<div class="note">' + echapper(c.note) + '</div>' : '') + '</div>').join('');

  const codes = new Set(a.devises.map(d => d.code));
  [...etat.devises].forEach(c => { if (!codes.has(c)) etat.devises.delete(c); });
  $('puces-devises').innerHTML = a.devises.map(d =>
    '<button class="puce" data-devise="' + echapper(d.code) + '" aria-pressed="' + etat.devises.has(d.code) + '">'
    + echapper(d.code) + '<span class="n chiffres">' + nombreFr(d.lignes) + '</span></button>').join('');
  rendrePucesColonnes();
}

function champHTML(c, s){
  const absent = c.statut === 'absent';
  const icone = absent ? ICONES.ko : (c.statut === 'neutre' ? ICONES.neutre : ICONES.ok);
  const vide = c.cle === 'devise' ? 'Aucune, tout en XOF' : 'Aucune colonne';
  let zone = '<button class="choix" data-cle="' + c.choix + '" aria-haspopup="listbox" aria-label="Colonne pour ' + echapper(c.libelle) + '">'
    + (c.colonne ? '<span>' + echapper(c.colonne) + '</span>' : '<span class="aucune">' + vide + '</span>')
    + ICONES.chevrons + '</button>';
  // En montant unique : le bouton rond montre ce qui part de ce côté
  // (+ ou −, ou D / C si le fichier porte une colonne de sens) ; un appui inverse.
  if (s.mode === 'montant' && s.montant && (c.cle === 'debit' || c.cle === 'credit')){
    const cote = (c.cle === 'debit') === !s.inverse;
    const marque = s.lettres ? (cote ? 'D' : 'C') : (cote ? '+' : '\u2212');
    zone += '<button class="signe" data-inverser title="Inverser débit et crédit" aria-label="Inverser débit et crédit">' + marque + '</button>';
  }
  const manque = c.manquantes === null ? '<span></span>'
    : '<span class="manque chiffres">valeurs manquantes : ' + nombreFr(c.manquantes) + ' sur ' + nombreFr(c.total) + '</span>';
  return '<div class="champ' + (absent ? ' absent' : '') + '">' + icone
    + '<span class="nom-champ">' + echapper(c.libelle) + '</span><div class="choix-zone">' + zone + '</div>' + manque + '</div>';
}

function rendreCorrespondance(){
  const s = structureActive();
  const par = cle => s.champs.find(c => c.cle === cle);
  const derive = s.mode === 'montant';
  const rappel = s.manquants.length
    ? '<div class="rappel">' + ICONES.ko + s.manquants.length + ' champ' + (s.manquants.length > 1 ? 's' : '')
      + ' requis absent' + (s.manquants.length > 1 ? 's' : '') + ' : ' + echapper(s.manquants.join(', ')) + '</div>'
    : '';
  $('correspondance').innerHTML = rappel
    + '<div class="groupe">' + champHTML(par('compte'), s) + '</div>'
    + '<div class="bascule"><div class="segment" role="group" aria-label="Forme des montants">'
    + '<button data-mode="dc" aria-pressed="' + !derive + '">Débit et crédit</button>'
    + '<button data-mode="montant" aria-pressed="' + derive + '">Montant unique</button></div></div>'
    + '<div class="groupe">' + champHTML(par('debit'), s) + champHTML(par('credit'), s) + '</div>'
    + '<div class="groupe">' + ['piece','libelle','impute','autorise','devise'].map(k => champHTML(par(k), s)).join('') + '</div>';
}

$('correspondance').addEventListener('click', e => {
  const mode = e.target.closest('[data-mode]'), choix = e.target.closest('[data-cle]'), signe = e.target.closest('[data-inverser]');
  if (signe) corriger({inverse: !structureActive().inverse});
  else if (mode && mode.getAttribute('aria-pressed') !== 'true') corriger({mode: mode.dataset.mode});
  else if (choix) ouvrirChoix(choix, choix.dataset.cle);
});
$('onglets-structures').addEventListener('click', e => {
  const b = e.target.closest('[data-structure]');
  if (!b) return;
  etat.structure = b.dataset.structure;
  rendreAnalyse(etat.analyse);
});

/* ---------------- choix d'une colonne, avec aperçu de ses valeurs ---------------- */
function fermerChoix(){ document.querySelectorAll('.voile,.popover').forEach(x => x.remove()); }

function ouvrirChoix(bouton, cle){
  fermerChoix();
  const s = structureActive();
  const champ = s.champs.find(c => c.choix === cle);
  const actuelle = champ ? champ.colonne : null;
  const vide = cle === 'devise' ? 'Aucune, tout en XOF' : 'Aucune colonne';

  const voile = document.createElement('div'); voile.className = 'voile';
  const pop = document.createElement('div'); pop.className = 'popover'; pop.setAttribute('role', 'dialog');
  pop.innerHTML = '<div class="titre-pop">Colonne pour <strong>' + echapper(NOMS[cle] || cle) + '</strong></div>'
    + '<input type="search" placeholder="Rechercher une colonne" aria-label="Rechercher une colonne">'
    + '<div class="options" role="listbox"></div>';
  const liste = pop.querySelector('.options'), recherche = pop.querySelector('input');

  function rendreOptions(filtre){
    const f = (filtre || '').toLowerCase();
    let html = f ? '' : '<button class="option" data-col="" aria-selected="' + (!actuelle) + '"><span class="coche">' + ICONES.coche + '</span><span class="aucune">' + vide + '</span></button>';
    s.colonnes.filter(c => !f || c.nom.toLowerCase().includes(f) || c.exemples.some(x => x.toLowerCase().includes(f))).forEach(c => {
      html += '<button class="option" role="option" data-col="' + echapper(c.nom) + '" aria-selected="' + (c.nom === actuelle) + '">'
        + '<span class="coche">' + ICONES.coche + '</span><span><span class="nom-col">' + echapper(c.nom) + '</span>'
        + '<span class="ex">' + (c.exemples.length ? c.exemples.map(x => '<span>' + echapper(x) + '</span>').join('') : '<span>vide</span>') + '</span>'
        + '</span></button>';
    });
    liste.innerHTML = html || '<div class="titre-pop">Aucune colonne ne correspond.</div>';
  }
  rendreOptions('');
  document.body.appendChild(voile); document.body.appendChild(pop);

  const r = bouton.getBoundingClientRect(), largeur = pop.offsetWidth, hauteur = pop.offsetHeight;
  const gauche = Math.min(Math.max(12, r.left), window.innerWidth - largeur - 12);
  let haut = r.bottom + 6;
  if (haut + hauteur > window.innerHeight - 12) haut = Math.max(12, r.top - hauteur - 6);
  pop.style.left = gauche + 'px'; pop.style.top = haut + 'px';
  recherche.focus();

  voile.addEventListener('click', fermerChoix);
  recherche.addEventListener('input', () => rendreOptions(recherche.value));
  recherche.addEventListener('keydown', e => { if (e.key === 'Enter'){ const o = liste.querySelector('.option'); if (o) o.click(); } });
  pop.addEventListener('keydown', e => { if (e.key === 'Escape'){ fermerChoix(); bouton.focus(); } });
  liste.addEventListener('click', e => {
    const o = e.target.closest('.option');
    if (!o) return;
    fermerChoix();
    corriger({cle: cle, colonne: o.dataset.col || null});
  });
}

/* ---------------- extraction ---------------- */
function colonnesDisponibles(){
  return etat.extraction ? etat.extraction.colonnes : (etat.analyse ? etat.analyse.colonnes : []);
}
function colonnesAffichees(){
  const dispo = colonnesDisponibles();
  if (etat.colonnesChoisies) return dispo.filter(c => etat.colonnesChoisies.has(c));
  return dispo.filter(c => c !== 'Fichier' || ids().length > 1);
}
function rendrePucesColonnes(){
  const affichees = new Set(colonnesAffichees());
  $('puces-colonnes').innerHTML = colonnesDisponibles().map(c =>
    '<button class="puce" data-colonne="' + echapper(c) + '" aria-pressed="' + affichees.has(c) + '">' + echapper(c) + '</button>').join('');
}
$('puces-colonnes').addEventListener('click', e => {
  const b = e.target.closest('[data-colonne]');
  if (!b) return;
  const courantes = new Set(colonnesAffichees());
  courantes.has(b.dataset.colonne) ? courantes.delete(b.dataset.colonne) : courantes.add(b.dataset.colonne);
  if (!courantes.size) return;
  etat.colonnesChoisies = courantes;
  rendrePucesColonnes();
  if (etat.extraction) preparerTableau();
});
$('puces-devises').addEventListener('click', e => {
  const b = e.target.closest('[data-devise]');
  if (!b) return;
  const c = b.dataset.devise;
  etat.devises.has(c) ? etat.devises.delete(c) : etat.devises.add(c);
  b.setAttribute('aria-pressed', etat.devises.has(c));
});

const CHAMPS_FILTRE = {compte:'f-compte', montant:'f-montant', piece:'f-piece', libelle:'f-libelle',
                       impute:'f-impute', autorise:'f-autorise', date_transaction:'f-date-transaction', date_valeur:'f-date-valeur'};
function filtres(){
  const f = {devises: [...etat.devises].sort()};
  Object.entries(CHAMPS_FILTRE).forEach(([cle, id]) => f[cle] = $(id).value);
  return f;
}
function viderExtraction(){
  etat.extraction = null;
  $('cadre').classList.add('cache');
  $('bilan').textContent = ''; $('criteres').innerHTML = '';
  $('btn-excel').disabled = true; $('btn-csv').disabled = true;
}

$('btn-extraire').addEventListener('click', async () => {
  if (!ids().length) return;
  erreur(''); $('bilan').textContent = 'Extraction…';
  try {
    const f = filtres();
    const d = await api('/api/extraire', {fichiers: ids(), filtres: f});
    etat.extraction = {total: d.total, colonnes: d.colonnes, blocs: new Map([[0, d.lignes]]), attente: new Set(),
                       filtres: f, fichiers: ids()};
    $('bilan').textContent = nombreFr(d.total) + ' ligne' + (d.total > 1 ? 's' : '');
    $('criteres').innerHTML = (d.criteres.length ? d.criteres : ['aucun filtre, tout le périmètre']).map(c => '<span class="critere">' + echapper(c) + '</span>').join('');
    $('btn-excel').disabled = d.total === 0; $('btn-csv').disabled = d.total === 0;
    rendrePucesColonnes();
    preparerTableau();
  } catch (err) { erreur(err.message); $('bilan').textContent = ''; }
});
document.querySelectorAll('.grille input').forEach(i => i.addEventListener('keydown', e => { if (e.key === 'Enter') $('btn-extraire').click(); }));

function preparerTableau(){
  const ex = etat.extraction, cols = colonnesAffichees();
  $('cadre').classList.remove('cache');
  $('cadre').scrollTop = 0;
  // Le libellé prend la place restante ; sans libellé, le tableau reste compact.
  const somme = cols.filter(c => c !== 'Libellé').reduce((a, c) => a + (LARGEURS[c] || 140), 0);
  const tableau = $('cadre').querySelector('table');
  if (cols.includes('Libellé')){ tableau.style.width = '100%'; tableau.style.minWidth = (somme + 200) + 'px'; }
  else { tableau.style.width = somme + 'px'; tableau.style.minWidth = '0'; }
  $('cols').innerHTML = cols.map(c => c === 'Libellé' ? '<col>' : '<col style="width:' + (LARGEURS[c] || 140) + 'px">').join('');
  $('entete').innerHTML = '<tr>' + cols.map(c => '<th class="' + (NUMERIQUES.has(c) ? 'nombre' : '') + '">' + echapper(c) + '</th>').join('') + '</tr>';
  ex.index = cols.map(c => ex.colonnes.indexOf(c));
  ex.vue = cols;
  rendreFenetre();
}

function rendreFenetre(){
  const ex = etat.extraction;
  if (!ex) return;
  const total = ex.total, cadre = $('cadre');
  if (!total){ $('corps').innerHTML = '<tr><td colspan="' + ex.vue.length + '" style="color:var(--second)">Aucune ligne ne correspond.</td></tr>'; return; }
  // Zone visible des lignes : le cadre, moins l'en-tête qui reste collé en haut.
  const h = ex.hauteur || HAUTEUR, vue = cadre.clientHeight - ($('entete').offsetHeight || 0);
  const reelle = total * h, virtuelle = Math.min(reelle, window.__hauteurMax || HAUTEUR_MAX);
  // Au-delà de la hauteur maximale, le défilement devient proportionnel :
  // le bas de la barre correspond toujours à la dernière ligne.
  const position = reelle > virtuelle
    ? Math.min(reelle - vue, cadre.scrollTop / Math.max(1, virtuelle - vue) * Math.max(0, reelle - vue))
    : cadre.scrollTop;
  const visible = Math.min(total - 1, Math.floor(position / h));
  const premier = Math.min(visible, Math.max(0, visible - 12, Math.ceil((position - cadre.scrollTop) / h)));
  const dernier = Math.min(total, visible + Math.ceil(vue / h) + 14);
  const haut = Math.max(0, cadre.scrollTop - (position - premier * h));
  const bas = Math.max(0, virtuelle - haut - (dernier - premier) * h);
  for (let b = Math.floor(premier / BLOC); b <= Math.floor((dernier - 1) / BLOC); b++) chargerBloc(b);
  const n = ex.vue.length;
  let html = '<tr class="espace"><td colspan="' + n + '" style="height:' + haut + 'px"></td></tr>';
  for (let i = premier; i < dernier; i++){
    const bloc = ex.blocs.get(Math.floor(i / BLOC));
    const ligne = bloc ? bloc[i % BLOC] : null;
    if (!ligne){ html += '<tr class="attente"><td colspan="' + n + '">Chargement…</td></tr>'; continue; }
    html += '<tr' + (i % 2 ? ' class="pair"' : '') + '>' + ex.index.map((k, j) => {
      const v = ligne[k], nom = ex.vue[j], classe = NUMERIQUES.has(nom) ? ' class="nombre"' : '';
      if (v === null || v === '') return '<td' + classe + '></td>';
      const texte = typeof v === 'number' ? nombreFr(v, Number.isInteger(v) ? 0 : 2) : String(v);
      return '<td' + classe + ' title="' + echapper(texte) + '">' + echapper(texte) + '</td>';
    }).join('') + '</tr>';
  }
  html += '<tr class="espace"><td colspan="' + n + '" style="height:' + bas + 'px"></td></tr>';
  $('corps').innerHTML = html;
  // Hauteur réelle d'une ligne, mesurée une fois : sinon l'écart s'accumule
  // sur des centaines de milliers de lignes.
  if (!ex.hauteur){
    const reelle = $('corps').querySelector('tr:not(.espace):not(.attente)');
    if (reelle){ ex.hauteur = reelle.getBoundingClientRect().height || HAUTEUR; if (Math.abs(ex.hauteur - HAUTEUR) > .05) rendreFenetre(); }
  }
}

async function chargerBloc(b){
  const ex = etat.extraction;
  if (ex.blocs.has(b) || ex.attente.has(b)) return;
  ex.attente.add(b);
  try {
    const d = await api('/api/lignes', {debut: b * BLOC, fin: (b + 1) * BLOC});
    if (etat.extraction !== ex) return;
    ex.blocs.set(b, d.lignes);
    requestAnimationFrame(rendreFenetre);
  } catch (err) { erreur(err.message); }
  finally { ex.attente.delete(b); }
}
let rafDefilement = null;
$('cadre').addEventListener('scroll', () => {
  if (rafDefilement) return;
  rafDefilement = requestAnimationFrame(() => { rafDefilement = null; rendreFenetre(); });
});

$('btn-effacer').addEventListener('click', () => {
  Object.values(CHAMPS_FILTRE).forEach(id => $(id).value = '');
  etat.devises.clear();
  document.querySelectorAll('[data-devise]').forEach(b => b.setAttribute('aria-pressed', 'false'));
  viderExtraction();
});

async function exporter(format, bouton){
  const libelle = bouton.textContent;
  bouton.disabled = true; bouton.textContent = 'Préparation…';
  try {
    const r = await fetch('/api/exporter', {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({fichiers: etat.extraction.fichiers, filtres: etat.extraction.filtres,
                            colonnes: colonnesAffichees(), format: format})});
    if (!r.ok){ const d = await r.json().catch(() => ({})); throw new Error(d.erreur || 'Export impossible.'); }
    const blob = await r.blob(), lien = document.createElement('a');
    lien.href = URL.createObjectURL(blob);
    lien.download = 'extraction.' + (format === 'csv' ? 'csv' : 'xlsx');
    lien.click();
    setTimeout(() => URL.revokeObjectURL(lien.href), 4000);
  } catch (err) { erreur(err.message); }
  finally { bouton.disabled = false; bouton.textContent = libelle; }
}
(async () => {
  try {
    const d = await (await fetch('/api/fichiers')).json();
    if (d.fichiers && d.fichiers.length){
      etat.fichiers = d.fichiers;
      d.fichiers.forEach(f => etat.selection.add(f.id));
      rendreFichiers(); analyser();
    }
  } catch (e) {}
})();
$('btn-excel').addEventListener('click', e => exporter('xlsx', e.currentTarget));
$('btn-csv').addEventListener('click', e => exporter('csv', e.currentTarget));
</script>
</body>
</html>
"""


# =====================================================================
#  9. Démarrage
# =====================================================================

if __name__ == "__main__":
    print("\n  Grand livre : http://127.0.0.1:5000")
    print("  Arrêt : Ctrl + C\n")
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
