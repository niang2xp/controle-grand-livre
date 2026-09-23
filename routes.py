# -*- coding: utf-8 -*-
"""
routes.py — Les adresses que l'interface appelle.

Chaque route lit la demande, appelle la fonction correspondante de
services.py et renvoie la réponse. Aucune logique métier ici.
"""

import os

from flask import Blueprint, jsonify, request, send_file

import services

routes = Blueprint("routes", __name__)

INTERFACE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "interface.html")


def corps():
    return request.get_json(silent=True) or {}


@routes.errorhandler(services.ErreurService)
def erreur_service(err):
    return jsonify({"erreur": str(err)}), 400


@routes.route("/")
def accueil():
    return send_file(INTERFACE, mimetype="text/html; charset=utf-8")


@routes.route("/api/fichiers", methods=["POST"])
def ajouter_fichier():
    # Le fichier arrive brut, sans formulaire : il est écrit sur disque par blocs.
    fichier = services.enregistrer_fichier(request.args.get("nom"), request.stream)
    return jsonify({"fichier": fichier})


@routes.route("/api/fichiers", methods=["GET"])
def lister_fichiers():
    return jsonify({"fichiers": services.lister_fichiers()})


@routes.route("/api/fichiers/<fid>", methods=["DELETE"])
def retirer_fichier(fid):
    services.retirer_fichier(fid)
    return jsonify({"ok": True})


@routes.route("/api/feuille", methods=["POST"])
def changer_feuille():
    d = corps()
    return jsonify({"fichier": services.changer_feuille(d.get("id"), d.get("feuille"))})


@routes.route("/api/analyse", methods=["POST"])
def analyse():
    return jsonify(services.analyse(corps().get("fichiers", [])))


@routes.route("/api/correspondance", methods=["POST"])
def correspondance():
    d = corps()
    return jsonify(services.corriger(d.get("fichiers", []), d.get("signature"),
                                     cle=d.get("cle"), colonne=d.get("colonne"),
                                     mode=d.get("mode"), inverse=d.get("inverse")))


@routes.route("/api/extraire", methods=["POST"])
def extraire():
    d = corps()
    return jsonify(services.extraire(d.get("fichiers", []), d.get("filtres")))


@routes.route("/api/lignes", methods=["POST"])
def lignes():
    d = corps()
    return jsonify(services.lignes(d.get("debut", 0), d.get("fin")))


@routes.route("/api/exporter", methods=["POST"])
def exporter():
    d = corps()
    contenu, nom, type_mime = services.exporter(d.get("fichiers", []), d.get("filtres"),
                                                d.get("colonnes"), d.get("format"))
    return send_file(contenu, mimetype=type_mime, as_attachment=True, download_name=nom)
