"""Paramètres du flash quotidien : seuils, périmètre, secteurs, interlocuteurs.

Tout ce qui peut changer sans toucher au code se règle ici.
"""
from __future__ import annotations

import datetime as dt

# ------------------------------------------------------------------ seuils
SEUILS = {
    "ca_vs_n1": -0.10,        # alerte si CA vs N-1 < -10 %
    "taux_pts": -3.0,         # alerte si écart de taux de marge < -3 pts
    "marge_vs_n1": -0.10,     # alerte si marge en valeur vs N-1 < -10 %
    "debit_vs_n1": -0.05,     # trafic en baisse si débit < -5 %
    "panier_vs_n1": -0.05,    # offre en baisse si panier < -5 %
    "ca_n1_min": 500_000,     # CA N-1 minimum d'un rayon x site pour alerter (FCFA)
    "taux_extreme_pts": -10.0,  # au-delà : on soupçonne une erreur de prix ou un déstockage
}

# ------------------------------------------------------------------ secteurs
DEPARTEMENTS = {
    "01 - PGC": "PGC",
    "02 - PRODUITS FRAIS": "Frais",
    "03 - BAZAR": "Non-alimentaire",
    "04 - EPCS": "Non-alimentaire",
    "06 - TEXTILE": "Non-alimentaire",
}
SECTEURS = ["PGC", "Frais", "Non-alimentaire"]

# ------------------------------------------------------------------ sites
FORMATS = ["Hyper", "Market", "Market ex-Supeco", "Supeco"]
SITES = [
    # site (libellé Power BI), format, zone de pricing, date de conversion (None si non converti)
    ("10301 - Hyper Marcory", "Hyper", "AB", None),
    ("10202 - Hyper Palmeraie", "Hyper", "AB", None),
    ("10203 - Hyper Yopougon", "Hyper", "CD", None),
    ("10206 - Market Kokoh Mall", "Market", "AB", None),
    ("10208 - Market Riviera", "Market", "AB", None),
    ("10209 - Market 2 Plateaux", "Market", "AB", None),
    ("10705 - Market 7 Décembre", "Market", "AB", None),
    ("10604 - Market Cité verte", "Market ex-Supeco", "CD", dt.date(2026, 4, 1)),
    ("10605 - Market Aboboté", "Market ex-Supeco", "CD", dt.date(2026, 8, 26)),
    ("10603 - Market Toit rouge", "Market ex-Supeco", "CD", dt.date(2026, 10, 22)),
    ("10601 - Supeco Niangon", "Supeco", "CD", None),
    ("10602 - Supeco Terminus 47", "Supeco", "CD", None),
]

# ------------------------------------------------------------------ interlocuteurs
# Pôles de l'équipe category management, pricing et données
POLES = {"pricing": "Pricing", "bdd": "Base de données", "promo": "Promo"}

# Délai de réponse des retours attendus : lendemain 12h (lundi pour un flash du vendredi ou du samedi)
HEURE_REPONSE = "12h"

# Nombre de lignes affichées
TOP_ALERTES = 5
TOP_RAYONS = 5
MAX_RETOURS = 6
