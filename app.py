# -*- coding: utf-8 -*-
"""
app.py — Contrôle et extraction d'un grand livre
================================================

Lancement :
    pip install flask pandas openpyxl
    python app.py

Le serveur démarre sur http://127.0.0.1:5000 et ouvre le navigateur.

Ce que fait l'application :
  1. charge un fichier .xlsx / .xls / .csv / .txt ;
  2. repère les colonnes attendues quel que soit leur libellé d'origine
     (accents, majuscules, abréviations) ;
  3. signale en rouge les champs requis absents du fichier ;
  4. renomme les colonnes reconnues en libellés propres et stables ;
  5. contrôle l'état de chaque colonne (type, valeurs manquantes, doublons) ;
  6. extrait les lignes voulues (compte, pièce, libellé, date, montant)
     et les exporte en Excel ou CSV.
"""

import io
import math
import os
import re
import threading
import unicodedata
import uuid
import webbrowser
from datetime import date, datetime

import numpy as np
import pandas as pd
from flask import Flask, Response, jsonify, request, send_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 Mo

# Mémoire de travail du serveur : jeton -> fichier chargé.
# L'application est mono-poste : quelques fichiers suffisent.
STORE = {}
STORE_MAX = 5


# =====================================================================
#  1. Normalisation des noms de colonnes
# =====================================================================

def normaliser(nom):
    """Réduit un nom de colonne à sa forme comparable.

    'N°compte'   -> 'n compte'
    'Lettrage '  -> 'lettrage'      (l'espace final disparaît)
    'libellés'   -> 'libelles'
    """
    texte = unicodedata.normalize("NFKD", str(nom))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    texte = texte.lower()
    texte = re.sub(r"[^a-z0-9]+", " ", texte)
    return re.sub(r"\s+", " ", texte).strip()


# Schéma des champs recherchés.
#   requis   : bloque le traitement s'il est absent
#   exact    : formes normalisées reconnues telles quelles (priorité 1)
#   contient : fragments recherchés dans le nom (priorité 2)
CHAMPS = [
    {
        "cle": "compte",
        "libelle": "Compte",
        "sortie": "Compte",
        "requis": True,
        "exact": ["n compte", "no compte", "num compte", "numero compte", "numero de compte",
                  "compte general", "compte", "cpte", "account", "account number", "gl account"],
        "contient": ["n compte", "numero compte", "compte", "cpte", "account"],
    },
    {
        "cle": "piece",
        "libelle": "Pièce",
        "sortie": "Pièce",
        "requis": True,
        "exact": ["n piece", "n pieces", "no piece", "num piece", "numero piece",
                  "numero de piece", "piece", "pieces", "document", "n document", "voucher"],
        "contient": ["piece", "voucher", "document"],
    },
    {
        "cle": "libelle",
        "libelle": "Libellé",
        "sortie": "Libellé",
        "requis": True,
        "exact": ["libelle", "libelles", "intitule", "designation", "description",
                  "objet", "narration", "text", "label"],
        "contient": ["libelle", "intitule", "designation", "description", "narration"],
    },
    {
        "cle": "montant",
        "libelle": "Montant",
        "sortie": "Montant",
        "requis": True,
        "exact": ["montant", "montant signe", "amount", "valeur", "value", "mnt"],
        "contient": ["montant", "amount"],
    },
    {
        "cle": "impute",
        "libelle": "Imputé par",
        "sortie": "Imputé par",
        "requis": True,
        "exact": ["impute", "impute par", "imputation", "saisi par", "saisie par",
                  "utilisateur", "user", "operateur", "auteur", "created by", "entered by",
                  "matricule saisie", "code utilisateur"],
        "contient": ["impute", "saisi", "utilisateur", "operateur", "auteur",
                     "created by", "entered by"],
    },
    {
        "cle": "autorise",
        "libelle": "Autorisé par",
        "sortie": "Autorisé par",
        "requis": True,
        "exact": ["autorise", "autorise par", "autorisation", "valide par", "validation",
                  "approuve par", "approbateur", "validateur", "approved by",
                  "authorized by", "authorised by", "checked by"],
        "contient": ["autorise", "authoris", "authoriz", "valide par", "validateur",
                     "approuve", "approved", "approbateur"],
    },
    # --- champs complémentaires : détectés, jamais bloquants -------------
    {
        "cle": "debit",
        "libelle": "Débit",
        "sortie": "Débit",
        "requis": False,
        "exact": ["debit", "montant debit", "dt"],
        "contient": ["debit"],
    },
    {
        "cle": "credit",
        "libelle": "Crédit",
        "sortie": "Crédit",
        "requis": False,
        "exact": ["credit", "montant credit", "ct"],
        "contient": ["credit"],
    },
    {
        "cle": "date",
        "libelle": "Date",
        "sortie": "Date",
        "requis": False,
        "exact": ["date", "date comptable", "date ecriture", "date piece",
                  "date operation", "posting date"],
        "contient": ["date"],
    },
    {
        "cle": "journal",
        "libelle": "Journal",
        "sortie": "Journal",
        "requis": False,
        "exact": ["journal", "journaux", "code journal", "code journaux", "jal", "jnl"],
        "contient": ["journal", "journaux"],
    },
    {
        "cle": "lettrage",
        "libelle": "Lettrage",
        "sortie": "Lettrage",
        "requis": False,
        "exact": ["lettrage", "lettre", "code lettrage", "matching"],
        "contient": ["lettrage", "matching"],
    },
    {
        "cle": "solde",
        "libelle": "Solde",
        "sortie": "Solde",
        "requis": False,
        "exact": ["solde", "solde cumule", "balance"],
        "contient": ["solde", "balance"],
    },
]

CHAMPS_PAR_CLE = {c["cle"]: c for c in CHAMPS}


def detecter_colonnes(colonnes):
    """Associe chaque champ du schéma à une colonne du fichier.

    Une colonne n'est attribuée qu'une seule fois. Les correspondances
    exactes passent avant les correspondances partielles.
    """
    normalisees = {col: normaliser(col) for col in colonnes}
    prises = set()
    mapping = {}

    # Passe 1 — correspondances exactes, pour tous les champs.
    # Elle doit précéder la passe partielle : sans cela, « Montant Débit »
    # serait capté par le champ Montant avant d'atteindre le champ Débit.
    for champ in CHAMPS:
        for col in colonnes:
            if col not in prises and normalisees[col] in champ["exact"]:
                prises.add(col)
                mapping[champ["cle"]] = col
                break

    # Passe 2 — correspondances partielles, du champ le plus spécifique
    # au plus général. Montant passe en dernier : c'est le terme le plus
    # susceptible d'apparaître dans l'intitulé d'une autre colonne.
    ordre_partiel = ["debit", "credit", "compte", "piece", "libelle", "impute",
                     "autorise", "journal", "lettrage", "date", "solde", "montant"]
    for cle in ordre_partiel:
        if cle in mapping:
            continue
        champ = CHAMPS_PAR_CLE[cle]
        for col in colonnes:
            if col not in prises and any(frag in normalisees[col] for frag in champ["contient"]):
                prises.add(col)
                mapping[cle] = col
                break

    return mapping


# =====================================================================
#  2. Lecture des fichiers
# =====================================================================

EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".tsv")


def lire_fichier(octets, nom, feuille=None):
    """Retourne (DataFrame, liste des feuilles, feuille lue)."""
    ext = os.path.splitext(nom)[1].lower()

    if ext in (".xlsx", ".xlsm", ".xls"):
        moteur = "openpyxl" if ext in (".xlsx", ".xlsm") else "xlrd"
        try:
            classeur = pd.ExcelFile(io.BytesIO(octets), engine=moteur)
        except ImportError:
            raise ValueError(
                "Le format .xls (ancien Excel) demande la bibliothèque xlrd. "
                "Installez-la avec « pip install xlrd », ou enregistrez le fichier en .xlsx."
            )
        feuilles = list(classeur.sheet_names)
        cible = feuille if feuille in feuilles else feuilles[0]
        df = classeur.parse(cible)
        return df, feuilles, cible

    if ext in (".csv", ".txt", ".tsv"):
        derniere = None
        for encodage in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(octets), sep=None, engine="python",
                                 encoding=encodage)
                return df, [], None
            except Exception as err:  # encodage ou séparateur inadapté
                derniere = err
        raise ValueError("Fichier illisible : séparateur ou encodage non reconnu (%s)." % derniere)

    raise ValueError("Format non pris en charge. Formats acceptés : %s." % ", ".join(EXTENSIONS))


# =====================================================================
#  3. Analyse du fichier
# =====================================================================

def valeur_json(valeur):
    """Convertit une valeur pandas en type acceptable par JSON."""
    if valeur is None:
        return None
    if isinstance(valeur, (pd.Timestamp, datetime, date)):
        if pd.isna(valeur):
            return None
        return valeur.strftime("%d/%m/%Y")
    # L'ordre compte : bool est une sous-classe de int en Python.
    if isinstance(valeur, (bool, np.bool_)):
        return bool(valeur)
    if isinstance(valeur, (int, np.integer)):
        return int(valeur)
    if isinstance(valeur, (float, np.floating)):
        nombre = float(valeur)
        return None if math.isnan(nombre) else nombre
    try:
        if pd.isna(valeur):
            return None
    except (TypeError, ValueError):
        pass
    return str(valeur)


def colonnes_affichage(normalise, origine_montant):
    """Colonnes montrées à l'écran et exportées.

    Un fichier porte soit une colonne de montant, soit le couple
    débit / crédit. On n'affiche jamais les deux.
    """
    if origine_montant == "colonne":
        ordre = ["Compte", "Montant"]
    else:
        ordre = ["Compte", "Débit", "Crédit"]
    ordre += ["Pièce", "Libellé", "Imputé par", "Autorisé par"]
    return [c for c in ordre if c in normalise.columns]


def construire_normalise(df, mapping, convention):
    """Construit le tableau aux colonnes renommées.

    convention : 'DC' pour Débit − Crédit, 'CD' pour Crédit − Débit.
    """
    sortie = pd.DataFrame(index=df.index)

    for champ in CHAMPS:
        cle = champ["cle"]
        if cle == "montant":
            continue
        if cle in mapping:
            serie = df[mapping[cle]]
            if cle == "compte":
                # Un numéro de compte est un identifiant, pas une grandeur :
                # le format texte préserve les zéros et permet le filtre par classe.
                serie = serie.apply(
                    lambda v: "" if pd.isna(v)
                    else (str(int(v)) if isinstance(v, (int, float, np.integer, np.floating))
                          and float(v).is_integer() else str(v).strip())
                )
            elif cle == "date":
                serie = pd.to_datetime(serie, errors="coerce")
            sortie[champ["sortie"]] = serie

    # Montant : colonne du fichier si elle existe, sinon reconstitution.
    origine_montant = None
    if "montant" in mapping:
        sortie["Montant"] = pd.to_numeric(df[mapping["montant"]], errors="coerce")
        origine_montant = "colonne"
    elif "debit" in mapping and "credit" in mapping:
        debit = pd.to_numeric(df[mapping["debit"]], errors="coerce").fillna(0)
        credit = pd.to_numeric(df[mapping["credit"]], errors="coerce").fillna(0)
        sortie["Montant"] = debit - credit if convention == "DC" else credit - debit
        origine_montant = "calcul"

    # Les colonnes non reconnues sont conservées telles quelles : aucune
    # information du fichier d'origine ne doit disparaître du tableau.
    utilisees = set(mapping.values())
    for col in df.columns:
        if col in utilisees:
            continue
        nom = str(col).strip() or "Colonne sans nom"
        while nom in sortie.columns:
            nom += " (bis)"
        sortie[nom] = df[col]

    # Ordre d'affichage lisible.
    ordre = ["Compte", "Date", "Journal", "Pièce", "Libellé", "Lettrage",
             "Débit", "Crédit", "Montant", "Solde", "Imputé par", "Autorisé par"]
    colonnes = [c for c in ordre if c in sortie.columns]
    colonnes += [c for c in sortie.columns if c not in colonnes]
    return sortie[colonnes], origine_montant


def analyser(df, mapping, convention, nom, feuilles, feuille):
    normalise, origine_montant = construire_normalise(df, mapping, convention)

    # --- état des champs attendus -------------------------------------
    total = int(len(df))

    # Le fichier porte soit un montant, soit le couple débit / crédit.
    # Sans colonne de montant, le couple débit / crédit est attendu en entier :
    # celui des deux qui manque doit être signalé, pas passé sous silence.
    cles_montant = ["montant"] if origine_montant == "colonne" else ["debit", "credit"]

    # Une écriture est au débit ou au crédit, jamais aux deux. N'est donc
    # réellement manquante que la ligne dépourvue des deux à la fois.
    manquantes_montant = None
    if "debit" in mapping and "credit" in mapping:
        manquantes_montant = int((df[mapping["debit"]].isna()
                                  & df[mapping["credit"]].isna()).sum())

    champs = []
    for cle in ["compte"] + cles_montant + ["piece", "libelle", "impute", "autorise"]:
        champ = CHAMPS_PAR_CLE[cle]
        entree = {"cle": cle, "libelle": champ["libelle"], "statut": "absent",
                  "manquantes": None, "doublons": None, "total": total}
        if cle in mapping:
            serie = df[mapping[cle]]
            entree["statut"] = "present"
            if cle in ("debit", "credit") and manquantes_montant is not None:
                entree["manquantes"] = manquantes_montant
            else:
                entree["manquantes"] = int(serie.isna().sum())
            entree["doublons"] = int(serie.dropna().duplicated().sum())
        champs.append(entree)

    manquants = [c["libelle"] for c in champs if c["statut"] == "absent"]

    # --- contrôles généraux -------------------------------------------
    controles = []

    def ajouter(intitule, valeur, niveau="info", note=None):
        controles.append({"intitule": intitule, "valeur": valeur,
                          "niveau": niveau, "note": note})

    ajouter("Volumétrie", "%s lignes · %s colonnes" % (f"{len(df):,}".replace(",", " "), len(df.columns)))

    doublons = int(df.duplicated().sum())
    ajouter("Lignes strictement identiques", str(doublons),
            "ok" if doublons == 0 else "alerte",
            None if doublons == 0 else "À examiner : double comptabilisation ou incident d'export.")

    vides = [str(c) for c in df.columns if df[c].isna().all()]
    if vides:
        ajouter("Colonnes entièrement vides", str(len(vides)), "attention",
                "Concernées : " + ", ".join(vides))

    if "Date" in normalise.columns and normalise["Date"].notna().any():
        debut = normalise["Date"].min()
        fin = normalise["Date"].max()
        illisibles = int(normalise["Date"].isna().sum())
        ajouter("Période couverte",
                "%s → %s" % (debut.strftime("%d/%m/%Y"), fin.strftime("%d/%m/%Y")),
                "ok" if illisibles == 0 else "attention",
                None if illisibles == 0 else "%s date(s) illisible(s) après conversion." % illisibles)

    if origine_montant == "colonne" and "Montant" in normalise.columns:
        ecart = float(pd.to_numeric(normalise["Montant"], errors="coerce").fillna(0).sum())
        ajouter("Équilibre général des montants", "écart de %s" % format_fr(ecart),
                "ok" if abs(ecart) < 0.005 else "alerte")
    elif "Débit" in normalise.columns and "Crédit" in normalise.columns:
        total_debit = float(pd.to_numeric(normalise["Débit"], errors="coerce").fillna(0).sum())
        total_credit = float(pd.to_numeric(normalise["Crédit"], errors="coerce").fillna(0).sum())
        ecart = total_debit - total_credit
        ajouter("Équilibre général débit / crédit", "écart de %s" % format_fr(ecart),
                "ok" if abs(ecart) < 0.005 else "alerte")

    return {
        "fichier": nom,
        "feuilles": feuilles,
        "feuille": feuille,
        "convention": convention,
        "lignes": int(len(df)),
        "colonnes_source": [str(c) for c in df.columns],
        "champs": champs,
        "manquants": manquants,
        "controles": controles,
        "origine_montant": origine_montant,
        "colonnes_normalisees": list(normalise.columns),
    }


def format_fr(nombre, decimales=0):
    try:
        texte = ("{:,.%df}" % decimales).format(float(nombre))
    except (TypeError, ValueError):
        return str(nombre)
    return texte.replace(",", " ").replace(".", ",")


def extraire_records(df):
    colonnes = [str(c) for c in df.columns]
    lignes = [[valeur_json(v) for v in ligne]
              for ligne in df.itertuples(index=False, name=None)]
    return colonnes, lignes


# =====================================================================
#  4. Routes
# =====================================================================

@app.route("/")
def accueil():
    return Response(PAGE, mimetype="text/html; charset=utf-8")


def ranger(jeton, entree):
    STORE[jeton] = entree
    while len(STORE) > STORE_MAX:
        STORE.pop(next(iter(STORE)))


@app.route("/api/charger", methods=["POST"])
def charger():
    fichier = request.files.get("fichier")
    if fichier is None or not fichier.filename:
        return jsonify({"erreur": "Aucun fichier reçu."}), 400

    nom = fichier.filename
    if not nom.lower().endswith(EXTENSIONS):
        return jsonify({"erreur": "Format non pris en charge. Formats acceptés : %s."
                                  % ", ".join(EXTENSIONS)}), 400

    octets = fichier.read()
    try:
        df, feuilles, feuille = lire_fichier(octets, nom)
    except ValueError as err:
        return jsonify({"erreur": str(err)}), 400
    except Exception as err:
        return jsonify({"erreur": "Lecture impossible : %s" % err}), 400

    if df.empty:
        return jsonify({"erreur": "Le fichier ne contient aucune ligne."}), 400

    mapping = detecter_colonnes(list(df.columns))
    convention = "DC"
    jeton = uuid.uuid4().hex
    ranger(jeton, {"octets": octets, "nom": nom, "df": df, "mapping": mapping,
                   "convention": convention, "feuilles": feuilles, "feuille": feuille})

    resultat = analyser(df, mapping, convention, nom, feuilles, feuille)
    resultat["jeton"] = jeton
    return jsonify(resultat)


@app.route("/api/reglage", methods=["POST"])
def reglage():
    """Change la feuille lue ou la convention de signe du montant."""
    donnees = request.get_json(silent=True) or {}
    entree = STORE.get(donnees.get("jeton"))
    if entree is None:
        return jsonify({"erreur": "Fichier expiré. Rechargez-le."}), 400

    if donnees.get("feuille") and donnees["feuille"] != entree["feuille"]:
        try:
            df, feuilles, feuille = lire_fichier(entree["octets"], entree["nom"],
                                                 donnees["feuille"])
        except Exception as err:
            return jsonify({"erreur": "Lecture de la feuille impossible : %s" % err}), 400
        entree.update({"df": df, "feuilles": feuilles, "feuille": feuille,
                       "mapping": detecter_colonnes(list(df.columns))})

    if donnees.get("convention") in ("DC", "CD"):
        entree["convention"] = donnees["convention"]

    resultat = analyser(entree["df"], entree["mapping"], entree["convention"],
                        entree["nom"], entree["feuilles"], entree["feuille"])
    resultat["jeton"] = donnees.get("jeton")
    return jsonify(resultat)


def appliquer_filtres(normalise, filtres):
    """Applique les critères d'extraction et retourne le sous-ensemble."""
    masque = pd.Series(True, index=normalise.index)
    appliques = []

    compte = (filtres.get("compte") or "").strip()
    if compte and "Compte" in normalise.columns:
        # Recherche par début de numéro : « 401100 » sort le compte,
        # « 7 » sort toute la classe.
        masque &= normalise["Compte"].astype(str).str.startswith(compte)
        appliques.append("compte commençant par %s" % compte)

    piece = (filtres.get("piece") or "").strip()
    if piece and "Pièce" in normalise.columns:
        serie = normalise["Pièce"]
        if pd.api.types.is_numeric_dtype(serie):
            # Évite l'écart entre « 797 » saisi et 797.0 stocké.
            cible = pd.to_numeric(piece, errors="coerce")
            masque &= (serie == cible) if pd.notna(cible) else False
        else:
            masque &= serie.astype(str).str.strip() == piece
        appliques.append("pièce n° %s" % piece)

    libelle = (filtres.get("libelle") or "").strip()
    if libelle and "Libellé" in normalise.columns:
        masque &= normalise["Libellé"].astype(str).str.contains(
            re.escape(libelle), case=False, na=False)
        appliques.append("libellé contenant « %s »" % libelle)

    for cle, colonne, etiquette in (("impute", "Imputé par", "imputé par"),
                                    ("autorise", "Autorisé par", "autorisé par")):
        valeur = (filtres.get(cle) or "").strip()
        if not valeur:
            continue
        if colonne not in normalise.columns:
            # Filtrer sur une colonne absente ne peut rien ramener : le dire
            # plutôt que d'ignorer le critère et de renvoyer tout le fichier.
            masque &= False
            appliques.append("%s : colonne absente du fichier" % etiquette)
            continue
        masque &= normalise[colonne].astype(str).str.contains(
            re.escape(valeur), case=False, na=False)
        appliques.append("%s contenant « %s »" % (etiquette, valeur))

    journal = (filtres.get("journal") or "").strip()
    if journal and "Journal" in normalise.columns:
        masque &= normalise["Journal"].astype(str).str.strip().str.upper() == journal.upper()
        appliques.append("journal %s" % journal.upper())

    if "Date" in normalise.columns:
        for cle, sens, etiquette in (("date_debut", "min", "à partir du"),
                                     ("date_fin", "max", "jusqu'au")):
            brut = (filtres.get(cle) or "").strip()
            if brut:
                borne = pd.to_datetime(brut, errors="coerce")
                if pd.notna(borne):
                    masque &= (normalise["Date"] >= borne) if sens == "min" else (normalise["Date"] <= borne)
                    appliques.append("%s %s" % (etiquette, borne.strftime("%d/%m/%Y")))

    if "Montant" in normalise.columns:
        for cle, sens, etiquette in (("montant_min", "min", "montant ≥"),
                                     ("montant_max", "max", "montant ≤")):
            brut = str(filtres.get(cle) or "").replace(" ", "").replace(",", ".").strip()
            if brut:
                try:
                    seuil = float(brut)
                except ValueError:
                    continue
                serie = normalise["Montant"]
                masque &= (serie >= seuil) if sens == "min" else (serie <= seuil)
                appliques.append("%s %s" % (etiquette,
                                            format_fr(seuil, 0 if seuil.is_integer() else 2)))

    return normalise[masque], appliques


@app.route("/api/extraire", methods=["POST"])
def extraire():
    donnees = request.get_json(silent=True) or {}
    entree = STORE.get(donnees.get("jeton"))
    if entree is None:
        return jsonify({"erreur": "Fichier expiré. Rechargez-le."}), 400

    normalise, origine = construire_normalise(entree["df"], entree["mapping"], entree["convention"])
    resultat, appliques = appliquer_filtres(normalise, donnees.get("filtres", {}))
    resultat = resultat[colonnes_affichage(normalise, origine)]

    colonnes, lignes = extraire_records(resultat)

    return jsonify({
        "total": int(len(resultat)),
        "criteres": appliques,
        "colonnes": colonnes,
        "lignes": lignes,
    })


@app.route("/api/exporter", methods=["POST"])
def exporter():
    donnees = request.get_json(silent=True) or {}
    entree = STORE.get(donnees.get("jeton"))
    if entree is None:
        return jsonify({"erreur": "Fichier expiré. Rechargez-le."}), 400

    normalise, origine = construire_normalise(entree["df"], entree["mapping"], entree["convention"])
    resultat, _ = appliquer_filtres(normalise, donnees.get("filtres", {}))
    resultat = resultat[colonnes_affichage(normalise, origine)]
    racine = os.path.splitext(entree["nom"])[0]

    if donnees.get("format") == "csv":
        tampon = io.BytesIO()
        tampon.write(resultat.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"))
        tampon.seek(0)
        return send_file(tampon, mimetype="text/csv", as_attachment=True,
                         download_name="extraction_%s.csv" % racine)

    tampon = io.BytesIO()
    with pd.ExcelWriter(tampon, engine="openpyxl") as writer:
        resultat.to_excel(writer, sheet_name="Extraction", index=False)
    tampon.seek(0)
    return send_file(
        tampon,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name="extraction_%s.xlsx" % racine)


# =====================================================================
#  5. Interface
# =====================================================================

PAGE = r"""<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Contrôle du grand livre</title>
<style>
  :root{
    --encre:#000000;
    --encre-clair:#5A5A5A;
    --page:#FFFFFF;
    --surface:#FFFFFF;
    --filet:#D6D6D6;
    --registre:#F6F6F6;
    --valide:#1D6B48;
    --absent:#B3261E;
    --absent-fond:#FDECEA;
    --absent-filet:#EFB4AE;
  }
  *{box-sizing:border-box}
  body{
    margin:0; background:var(--page); color:var(--encre);
    font-family:ui-sans-serif,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
    font-size:15px; line-height:1.5;
  }
  .mono{font-family:ui-monospace,SFMono-Regular,"Cascadia Mono","JetBrains Mono",Consolas,monospace;
        font-variant-numeric:tabular-nums}
  h1,h2,h3{margin:0; font-weight:600; letter-spacing:-.01em}
  h1{font-size:19px}
  h2{font-size:17px}
  h3{font-size:14px; color:var(--encre-clair)}
  p{margin:0}
  button{font:inherit; cursor:pointer}
  input,select{font:inherit; color:inherit}
  input:focus-visible,select:focus-visible,button:focus-visible{
    outline:2px solid var(--encre); outline-offset:2px}

  header{background:var(--encre); color:#fff; padding:13px 28px;
         display:flex; align-items:baseline; gap:18px; flex-wrap:wrap}
  header .fichier{font-size:13px; color:#D2D2D2}

  main{max-width:1180px; margin:0 auto; padding:28px 24px 72px; background:var(--page)}
  section{background:var(--surface); border:1px solid var(--filet);
          border-radius:4px; margin-bottom:20px}
  .tete{padding:16px 20px; border-bottom:1px solid var(--filet);
        display:flex; align-items:center; gap:14px; flex-wrap:wrap}
  .tete p{font-size:13px; color:var(--encre-clair)}
  .corps{padding:20px}

  /* dépôt du fichier */
  .depot{border:2px dashed var(--filet); border-radius:4px; background:var(--surface);
         padding:44px 24px; text-align:center; transition:border-color .15s, background .15s}
  .depot.survol{border-color:var(--encre); background:var(--registre)}
  .depot h2{margin-bottom:6px}
  .depot p{color:var(--encre-clair); font-size:14px; margin-bottom:18px}
  .bouton{background:var(--encre); color:#fff; border:1px solid var(--encre);
          border-radius:3px; padding:9px 18px; font-size:14px; font-weight:500}
  .bouton:hover{background:#2B2B2B}
  .bouton.secondaire{background:transparent; color:var(--encre)}
  .bouton.secondaire:hover{background:var(--page)}
  .bouton:disabled{opacity:.45; cursor:default}

  /* état des champs */
  .champ{display:grid; grid-template-columns:26px 1fr; gap:12px;
         padding:14px 0; border-top:1px solid var(--filet)}
  .champ:first-child{border-top:0}
  .champ .marque{padding-top:2px}
  .champ .titre strong{font-size:15px}
  .champ .mesures{font-size:13px; color:var(--encre-clair); margin-top:3px;
                  display:flex; gap:22px; flex-wrap:wrap}
  .champ.absent{background:var(--absent-fond); border-left:3px solid var(--absent);
                padding-left:14px; padding-right:14px; border-radius:3px; margin:8px 0}
  .champ.absent .titre strong{color:var(--absent)}
  .champ.absent .role{color:#7C2019}
  .marque svg{display:block}

  .rappel{display:flex; gap:12px; align-items:flex-start; background:var(--absent-fond);
          border:1px solid var(--absent-filet); border-radius:4px; padding:14px 16px;
          margin-bottom:18px}
  .rappel strong{color:var(--absent)}
  .rappel p{font-size:13px; color:#7C2019; margin-top:2px}

  /* contrôles */
  .controles{display:grid; grid-template-columns:repeat(auto-fill,minmax(250px,1fr)); gap:12px}
  .controle{background:var(--surface); border:1px solid var(--filet); border-radius:3px;
            padding:13px 15px}
  .controle .intitule{font-size:13px; color:var(--encre-clair)}
  .controle .valeur{font-size:15px; margin-top:2px}
  .controle .note{font-size:12px; color:var(--encre-clair); margin-top:5px}
  .controle.attention .valeur,
  .controle.alerte .valeur{color:var(--absent)}

  /* tableaux */
  .cadre{overflow:auto; border:1px solid var(--filet); border-radius:3px; max-height:460px}
  table{border-collapse:collapse; width:100%; font-size:13px}
  th,td{padding:7px 11px; text-align:left; white-space:nowrap;
        border-bottom:1px solid var(--filet)}
  thead th{position:sticky; top:0; background:var(--encre); color:#fff; font-weight:500;
           border-bottom:0; z-index:1}
  tbody tr:nth-child(even){background:var(--registre)}
  th.nombre,td.nombre{text-align:right}

  /* extraction */
  .grille{display:grid; grid-template-columns:repeat(auto-fill,minmax(200px,1fr)); gap:14px}
  label{display:block; font-size:13px; color:var(--encre-clair); margin-bottom:4px}
  input[type=text],input[type=date],select{
    width:100%; padding:7px 9px; border:1px solid var(--filet); border-radius:3px;
    background:var(--surface)}
  .actions{display:flex; gap:10px; margin-top:18px; flex-wrap:wrap; align-items:center}
  .bilan{font-size:13px; color:var(--encre-clair)}
  .critere{display:inline-block; background:var(--surface); border:1px solid var(--filet);
           border-radius:3px; padding:1px 8px; font-size:12px; margin:0 6px 6px 0}
  .erreur{background:var(--absent-fond); border:1px solid var(--absent-filet);
          color:#7C2019; padding:11px 14px; border-radius:3px; font-size:14px;
          margin-bottom:16px}
  .cache{display:none}
  .patiente{color:var(--encre-clair); font-size:14px}
  @media (max-width:640px){ main{padding:18px 14px 56px} .corps{padding:15px} }
</style>
</head>
<body>

<header>
  <span class="fichier" id="entete-fichier">aucun fichier chargé</span>
</header>

<main>
  <div id="zone-erreur"></div>

  <section id="bloc-depot">
    <div class="corps">
      <div class="depot" id="depot">
        <h2>Déposez l'export comptable</h2>
        <p>Excel (.xlsx, .xlsm, .xls), CSV ou texte délimité — 50 Mo maximum.</p>
        <button class="bouton" id="btn-parcourir">Choisir un fichier</button>
        <input type="file" id="input-fichier" class="cache" accept=".xlsx,.xlsm,.xls,.csv,.txt,.tsv">
        <p id="etat-depot" class="patiente" style="margin-top:16px"></p>
      </div>
    </div>
  </section>

  <div id="resultats" class="cache">

    <section>
      <div class="tete">
        <h2>Colonnes attendues</h2>
        <div id="choix-feuille" class="cache" style="margin-left:auto">
          <label for="feuille" style="margin:0 0 2px">Feuille</label>
          <select id="feuille"></select>
        </div>
      </div>
      <div class="corps">
        <div id="bloc-manquants"></div>
        <div id="liste-champs"></div>
      </div>
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
          <div>
            <label for="f-compte">Compte</label>
            <input type="text" id="f-compte" placeholder="401100">
          </div>
          <div>
            <label for="f-piece">Pièce</label>
            <input type="text" id="f-piece" placeholder="797">
          </div>
          <div>
            <label for="f-impute">Imputé par</label>
            <input type="text" id="f-impute" placeholder="nom de l'utilisateur">
          </div>
          <div>
            <label for="f-autorise">Autorisé par</label>
            <input type="text" id="f-autorise" placeholder="nom du validateur">
          </div>
          <div>
            <label for="f-journal">Journal</label>
            <input type="text" id="f-journal" placeholder="OD">
          </div>
          <div>
            <label for="f-libelle">Libellé contient</label>
            <input type="text" id="f-libelle" placeholder="provision">
          </div>
          <div>
            <label for="f-date-debut">Date à partir du</label>
            <input type="date" id="f-date-debut">
          </div>
          <div>
            <label for="f-date-fin">Date jusqu'au</label>
            <input type="date" id="f-date-fin">
          </div>
          <div>
            <label for="f-montant-min">Montant minimum</label>
            <input type="text" id="f-montant-min" placeholder="1000000">
          </div>
          <div>
            <label for="f-montant-max">Montant maximum</label>
            <input type="text" id="f-montant-max" placeholder="50000000">
          </div>
        </div>
        <div class="actions">
          <button class="bouton" id="btn-extraire">Extraire</button>
          <button class="bouton secondaire" id="btn-reinit">Tout effacer</button>
          <button class="bouton secondaire" id="btn-excel" disabled>Exporter en Excel</button>
          <button class="bouton secondaire" id="btn-csv" disabled>Exporter en CSV</button>
          <span class="bilan" id="bilan-extraction"></span>
        </div>
        <div id="criteres-appliques" style="margin-top:14px"></div>
        <div class="cadre" id="cadre-extraction" style="margin-top:6px"></div>
      </div>
    </section>

  </div>
</main>

<script>
const TRIANGLE = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#B3261E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.6 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.6a2 2 0 0 0-3.4 0Z"/><line x1="12" y1="9" x2="12" y2="13.5"/><line x1="12" y1="17.2" x2="12.01" y2="17.2"/></svg>';
const COCHE = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#1D6B48" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8 12.3 2.7 2.7L16 9.6"/></svg>';
let JETON = null;
let ANALYSE = null;
const $ = (id) => document.getElementById(id);

function echapper(t){
  return String(t).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function nombreFr(v, dec){
  return new Intl.NumberFormat('fr-FR', {minimumFractionDigits: dec||0, maximumFractionDigits: dec||0}).format(v);
}
function erreur(message){
  $('zone-erreur').innerHTML = message ? '<div class="erreur">' + echapper(message) + '</div>' : '';
  if (message) window.scrollTo({top:0, behavior:'smooth'});
}

/* ---------- chargement ---------- */
const depot = $('depot');
$('btn-parcourir').addEventListener('click', () => $('input-fichier').click());
$('input-fichier').addEventListener('change', e => { if (e.target.files[0]) envoyer(e.target.files[0]); });
['dragenter','dragover'].forEach(t => depot.addEventListener(t, e => {
  e.preventDefault(); depot.classList.add('survol');
}));
['dragleave','drop'].forEach(t => depot.addEventListener(t, e => {
  e.preventDefault(); depot.classList.remove('survol');
}));
depot.addEventListener('drop', e => { if (e.dataTransfer.files[0]) envoyer(e.dataTransfer.files[0]); });

async function envoyer(fichier){
  erreur('');
  $('etat-depot').textContent = 'Lecture de ' + fichier.name + '…';
  const corps = new FormData();
  corps.append('fichier', fichier);
  try{
    const reponse = await fetch('/api/charger', {method:'POST', body:corps});
    const donnees = await reponse.json();
    if (!reponse.ok){ erreur(donnees.erreur || 'Chargement impossible.'); $('etat-depot').textContent=''; return; }
    JETON = donnees.jeton;
    afficher(donnees);
    $('etat-depot').textContent = '';
  }catch(err){
    erreur('Le serveur ne répond pas : ' + err.message);
    $('etat-depot').textContent = '';
  }
}

async function reglage(champ, valeur){
  const charge = {jeton: JETON};
  charge[champ] = valeur;
  const reponse = await fetch('/api/reglage', {
    method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(charge)});
  const donnees = await reponse.json();
  if (!reponse.ok){ erreur(donnees.erreur || 'Réglage impossible.'); return; }
  afficher(donnees);
}
$('feuille').addEventListener('change', e => reglage('feuille', e.target.value));

/* ---------- restitution ---------- */
function afficher(a){
  ANALYSE = a;
  $('resultats').classList.remove('cache');
  $('bloc-depot').querySelector('.depot h2').textContent = 'Charger un autre fichier';
  $('entete-fichier').textContent = a.fichier + (a.feuille ? ' · feuille ' + a.feuille : '')
    + ' · ' + nombreFr(a.lignes) + ' lignes';

  if (a.feuilles && a.feuilles.length > 1){
    $('choix-feuille').classList.remove('cache');
    $('feuille').innerHTML = a.feuilles.map(f =>
      '<option value="' + echapper(f) + '"' + (f===a.feuille?' selected':'') + '>' + echapper(f) + '</option>').join('');
  } else {
    $('choix-feuille').classList.add('cache');
  }

  /* rappel des champs requis absents */
  if (a.manquants.length){
    $('bloc-manquants').innerHTML =
      '<div class="rappel">' + TRIANGLE +
      '<div><strong>' + a.manquants.length + ' champ' + (a.manquants.length>1?'s':'') +
      ' requis absent' + (a.manquants.length>1?'s':'') + ' : ' + echapper(a.manquants.join(', ')) +
      '</strong></div></div>';
  } else {
    $('bloc-manquants').innerHTML = '';
  }

  /* liste des champs */
  $('liste-champs').innerHTML = a.champs.map(c => {
    const absent = c.statut === 'absent';
    const detail = absent
      ? ''
      : '<div class="mesures"><span>valeurs manquantes : ' + nombreFr(c.manquantes)
        + ' sur ' + nombreFr(c.total) + '</span><span>doublons : '
        + nombreFr(c.doublons) + ' sur ' + nombreFr(c.total) + '</span></div>';
    return '<div class="champ' + (absent ? ' absent' : '') + '">'
      + '<div class="marque">' + (absent ? TRIANGLE : COCHE) + '</div><div>'
      + '<div class="titre"><strong>' + echapper(c.libelle) + '</strong></div>'
      + detail + '</div></div>';
  }).join('');

  /* contrôles généraux */
  $('liste-controles').innerHTML = a.controles.map(c =>
    '<div class="controle ' + c.niveau + '">'
    + '<div class="intitule">' + echapper(c.intitule) + '</div>'
    + '<div class="valeur mono">' + echapper(c.valeur) + '</div>'
    + (c.note ? '<div class="note">' + echapper(c.note) + '</div>' : '')
    + '</div>').join('');

  document.querySelector('#resultats section').scrollIntoView({behavior:'smooth', block:'start'});
}

function tableau(colonnes, lignes){
  if (!lignes.length) return '<p style="padding:16px" class="patiente">Aucune ligne.</p>';
  // Une colonne dont les valeurs sont des nombres est alignée à droite,
  // en-tête compris, pour que les chiffres se lisent en colonne.
  const nombre = colonnes.map((c, i) => lignes.some(l => typeof l[i] === 'number'));
  const entete = '<thead><tr>' + colonnes.map((c, i) =>
    '<th' + (nombre[i] ? ' class="nombre"' : '') + '>' + echapper(c) + '</th>').join('') + '</tr></thead>';
  const corps = '<tbody>' + lignes.map(l => '<tr>' + l.map((v, i) => {
    const classe = nombre[i] ? ' class="nombre mono"' : '';
    if (v === null || v === '') return '<td' + classe + '></td>';
    if (typeof v === 'number') return '<td' + classe + '>' + nombreFr(v, Number.isInteger(v)?0:2) + '</td>';
    return '<td' + classe + '>' + echapper(v) + '</td>';
  }).join('') + '</tr>').join('') + '</tbody>';
  return '<table>' + entete + corps + '</table>';
}

/* ---------- extraction ---------- */
function filtres(){
  return {
    compte: $('f-compte').value,
    piece: $('f-piece').value,
    impute: $('f-impute').value,
    autorise: $('f-autorise').value,
    journal: $('f-journal').value,
    libelle: $('f-libelle').value,
    date_debut: $('f-date-debut').value,
    date_fin: $('f-date-fin').value,
    montant_min: $('f-montant-min').value,
    montant_max: $('f-montant-max').value
  };
}

$('btn-extraire').addEventListener('click', async () => {
  if (!JETON) return;
  erreur('');
  $('bilan-extraction').textContent = 'Extraction…';
  const reponse = await fetch('/api/extraire', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({jeton: JETON, filtres: filtres()})});
  const d = await reponse.json();
  if (!reponse.ok){ erreur(d.erreur || 'Extraction impossible.'); $('bilan-extraction').textContent=''; return; }

  $('bilan-extraction').textContent = nombreFr(d.total) + ' ligne' + (d.total>1?'s':'');
  $('criteres-appliques').innerHTML = d.criteres.length
    ? d.criteres.map(c => '<span class="critere">' + echapper(c) + '</span>').join('')
    : '<span class="critere">aucun filtre — fichier complet</span>';
  $('cadre-extraction').innerHTML = tableau(d.colonnes, d.lignes);
  $('btn-excel').disabled = d.total === 0;
  $('btn-csv').disabled = d.total === 0;
});

$('btn-reinit').addEventListener('click', () => {
  ['f-compte','f-piece','f-impute','f-autorise','f-journal','f-libelle',
   'f-date-debut','f-date-fin','f-montant-min','f-montant-max'].forEach(id => $(id).value = '');
  $('cadre-extraction').innerHTML = '';
  $('criteres-appliques').innerHTML = '';
  $('bilan-extraction').textContent = '';
  $('btn-excel').disabled = true;
  $('btn-csv').disabled = true;
});

async function exporter(format){
  const reponse = await fetch('/api/exporter', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({jeton: JETON, filtres: filtres(), format: format})});
  if (!reponse.ok){
    const d = await reponse.json().catch(() => ({}));
    erreur(d.erreur || 'Export impossible.');
    return;
  }
  const blob = await reponse.blob();
  const lien = document.createElement('a');
  lien.href = URL.createObjectURL(blob);
  lien.download = 'extraction.' + (format === 'csv' ? 'csv' : 'xlsx');
  lien.click();
  URL.revokeObjectURL(lien.href);
}
$('btn-excel').addEventListener('click', () => exporter('xlsx'));
$('btn-csv').addEventListener('click', () => exporter('csv'));
</script>
</body>
</html>
"""


# =====================================================================
#  6. Démarrage
# =====================================================================

def ouvrir_navigateur():
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    print("\n  Contrôle du grand livre")
    print("  Serveur en écoute sur http://127.0.0.1:5000")
    print("  Arrêt : Ctrl + C\n")
    threading.Timer(1.2, ouvrir_navigateur).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
