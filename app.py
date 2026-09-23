# -*- coding: utf-8 -*-
"""
app.py — Outil de contrôle et d'extraction de grands livres.

Lancement :
    pip install -r requirements.txt
    python app.py

Le serveur démarre sur http://127.0.0.1:5000 et ouvre le navigateur.

Organisation du projet :
  app.py            démarrage du serveur (ce fichier)
  routes.py         les adresses appelées par l'interface
  services.py       la logique : lecture, correspondance, analyse, extraction, export
  interface.html    l'interface : HTML, CSS et JavaScript
  requirements.txt  les dépendances
"""

import socket
import sys
import threading
import webbrowser

from flask import Flask

import services
from routes import routes

app = Flask(__name__)
# Aucun plafond : la vraie limite est la mémoire de la machine.
app.config["MAX_CONTENT_LENGTH"] = None
app.register_blueprint(routes)


def deja_lancee(port):
    """Une instance répond-elle déjà sur ce port ?"""
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


if __name__ == "__main__":
    if deja_lancee(5000):
        # Lancée une seconde fois : on rouvre la page au lieu de planter,
        # et surtout sans effacer les fichiers de l'instance ouverte.
        print("\n  L'application tourne déjà : http://127.0.0.1:5000\n")
        webbrowser.open("http://127.0.0.1:5000")
        sys.exit(0)
    services.nettoyer_anciens_dossiers()
    print("\n  Grand livre : http://127.0.0.1:5000")
    print("  Arrêt : Ctrl + C\n")
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
