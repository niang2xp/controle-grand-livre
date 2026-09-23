# -*- coding: utf-8 -*-
"""
services.py — Toute la logique de l'outil, sans rien de web.

Sections :
  0. Dossier de travail et état en mémoire
  1. Outils de conversion (nombres, dates, identifiants, devises, sens)
  2. Détection proposée des colonnes
  3. Lecture et conversion des fichiers
  4. Table normalisée (compte, débit, crédit, pièce...)
  5. Analyse (colonnes attendues, état du fichier)
  6. Extraction (filtres)
  7. Fonctions appelées par les routes
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
from datetime import date, datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook


# =====================================================================
#  0. Dossier de travail et état en mémoire
# =====================================================================

PREFIXE_TRAVAIL = "grand_livre_"


def preparer_dossier_travail():
    """Dossier temporaire des fichiers déposés, effacé à l'arrêt."""
    dossier = tempfile.mkdtemp(prefix=PREFIXE_TRAVAIL)
    atexit.register(shutil.rmtree, dossier, ignore_errors=True)
    return dossier


DOSSIER_TRAVAIL = preparer_dossier_travail()


def nettoyer_anciens_dossiers():
    """Supprime les dossiers de travail laissés par une exécution interrompue
    brutalement : ils contiennent des données client qui n'ont pas à traîner
    sur le disque. Appelée au démarrage seulement quand aucune autre instance
    ne tourne, pour ne jamais effacer les fichiers d'une instance ouverte."""
    for ancien in glob.glob(os.path.join(tempfile.gettempdir(), PREFIXE_TRAVAIL + "*")):
        if os.path.abspath(ancien) != os.path.abspath(DOSSIER_TRAVAIL):
            shutil.rmtree(ancien, ignore_errors=True)

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
#  7. Fonctions appelées par les routes
# =====================================================================

class ErreurService(Exception):
    """Erreur destinée à l'utilisateur : son message s'affiche tel quel."""


def enregistrer_fichier(nom, flux):
    """Écrit sur disque un fichier reçu, le convertit, le lit, et le range."""
    nom = (nom or "fichier").strip()
    ext = os.path.splitext(nom)[1].lower()
    if ext not in EXTENSIONS:
        raise ErreurService("« %s » : format non pris en charge. Formats acceptés : %s."
                            % (nom, ", ".join(EXTENSIONS)))

    fid = uuid.uuid4().hex[:12]
    source = os.path.join(DOSSIER_TRAVAIL, fid + ext)
    # Écriture par blocs : un fichier volumineux ne passe jamais en entier
    # par la mémoire pendant l'envoi.
    taille = 0
    with open(source, "wb") as sortie:
        while True:
            bloc = flux.read(8 * 1024 * 1024)
            if not bloc:
                break
            sortie.write(bloc)
            taille += len(bloc)
    if taille == 0:
        os.remove(source)
        raise ErreurService("« %s » est vide." % nom)

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
        raise ErreurService("« %s » : %s" % (nom, err))
    except Exception as err:
        raise ErreurService("« %s » illisible : %s" % (nom, err))
    return infos_fichier(fichier)


def lister_fichiers():
    """Fichiers déjà chargés : la page les retrouve après un rechargement."""
    with VERROU:
        return [infos_fichier(f) for f in FICHIERS.values()]


def retirer_fichier(fid):
    with VERROU:
        fichier = FICHIERS.pop(fid, None)
        if fichier:
            for chemin in {fichier["source"], fichier.get("csv")}:
                if chemin and os.path.exists(chemin):
                    os.remove(chemin)
            RESULTAT.update({"cle": None, "table": None})


def changer_feuille(fid, feuille):
    fichier = FICHIERS.get(fid)
    if not fichier or not fichier["feuilles"]:
        raise ErreurService("Fichier introuvable.")
    try:
        with VERROU:
            fichier["feuilles"], fichier["feuille"] = convertir_excel(
                fichier["source"], fichier["csv"], feuille)
            charger(fichier)
    except Exception as err:
        raise ErreurService("Lecture de la feuille impossible : %s" % err)
    return infos_fichier(fichier)


def analyse(ids):
    with VERROU:
        return analyser(ids)


def corriger(ids, signature, cle=None, colonne=None, mode=None, inverse=None):
    """Change une correspondance, la forme des montants ou leur sens."""
    structure = STRUCTURES.get(signature)
    if structure is None:
        raise ErreurService("Structure inconnue.")
    with VERROU:
        # Vérification avant toute modification : une demande refusée ne
        # laisse rien de changé à moitié.
        if cle in CLES and colonne is not None and colonne not in structure["colonnes"]:
            raise ErreurService("Colonne inconnue : %s" % colonne)
        if mode in ("dc", "montant"):
            structure["mode"] = mode
        if isinstance(inverse, bool):
            structure["inverse"] = inverse
        if cle in CLES:
            structure["mapping"][cle] = colonne
        structure["version"] += 1
        return analyser(ids)


def extraire(ids, filtres):
    if not selection(ids):
        raise ErreurService("Aucun fichier sélectionné.")
    with VERROU:
        table, criteres = resultat(ids, filtres)
        return {"total": int(len(table)), "colonnes": list(table.columns),
                "criteres": criteres, "lignes": lignes_json(table, 0, 200)}


def lignes(debut, fin=None):
    """Un bloc de lignes du dernier résultat, pour l'affichage au défilement."""
    with VERROU:
        table = RESULTAT["table"]
        if table is None:
            raise ErreurService("Relancez l'extraction.")
        debut = max(0, int(debut or 0))
        fin = min(len(table), int(fin) if fin is not None else debut + 200)
        return {"debut": debut, "lignes": lignes_json(table, debut, fin)}


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


def exporter(ids, filtres, colonnes, format_export):
    """Retourne (contenu, nom du fichier, type) pour le téléchargement."""
    if not selection(ids):
        raise ErreurService("Aucun fichier sélectionné.")
    with VERROU:
        table, _ = resultat(ids, filtres, memoriser=False)
        choisies = [c for c in (colonnes or []) if c in table.columns] or list(table.columns)
        table = table[choisies].copy()
    # Montants entiers écrits sans décimale : 30000 et non 30000,0.
    for col in ("Débit", "Crédit"):
        if col in table.columns:
            valeurs = table[col].dropna()
            if len(valeurs) and bool((valeurs % 1 == 0).all()):
                table[col] = table[col].astype("Int64")

    if format_export == "csv":
        texte = table.to_csv(index=False, sep=";", decimal=",", date_format="%d/%m/%Y")
        return io.BytesIO(texte.encode("utf-8-sig")), "extraction.csv", "text/csv"

    if len(table) > 1048575:
        raise ErreurService("%s lignes : Excel ne peut pas en contenir plus de 1 048 575. "
                            "Exportez en CSV." % format_fr(len(table)))
    classeur = Workbook(write_only=True)
    feuille = classeur.create_sheet("Extraction")
    feuille.append(choisies)
    for ligne in table.itertuples(index=False, name=None):
        feuille.append([valeur_excel(v) for v in ligne])
    tampon = io.BytesIO()
    classeur.save(tampon)
    tampon.seek(0)
    return (tampon, "extraction.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
