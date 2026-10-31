# Flash quotidien · SmartBuyer Hub

Chaque matin : 3 extractions Power BI + le masque PowerPoint CFAO →
contrôle automatique → PPT de 6 slides pour les secteurs + Excel de détail (valeurs, sans formules).

## Installation

```
pip install -r requirements.txt
```

Organisation dans le hub :

```
modules/flash_quotidien/    moteur (config, calculs, PPT, Excel) et run_flash.py
pages/27_⚡_Flash_Quotidien.py
modules/flash_quotidien/assets/masque_cfao_consumer.pptx   facultatif, non versionné (dépôt public)
```

Lancer : `streamlit run app.py` (page « Flash Quotidien »), ou sans interface, depuis la racine du dépôt :

```
python -m modules.flash_quotidien.run_flash hier.xlsx semaine.xlsx sites.xlsx masque.pptx --date 2026-10-30 --out sorties
```

## Les 3 extractions attendues

Power BI, page des ventes CA, **affichage HT**. Le N-1 est le même jour de semaine de l'an dernier.
Toujours : menu « … » du visuel › Exporter les données › **Données avec la disposition actuelle** › .xlsx.

| # | Bouton période | Matrice | Colonnes indispensables | Pied de page attendu |
|---|---|---|---|---|
| 1 | Hier | Département › Rayon › Site, tout développé | Département, Rayon, Site, CA N-1, Budget, CA, Marge N-1, Marge, Débit N-1, Débit | Type d'affichage est Hier · Libelle est HT |
| 2 | Cette semaine | Même matrice que le 1 | Département, Rayon, Site, CA N-1, Budget, CA, Marge N-1, Marge | Type d'affichage est Cette Semaine · Libelle est HT |
| 3 | Hier | **Site uniquement** (une ligne par magasin) | Site, CA N-1, CA, Marge N-1, Marge, Débit N-1, Débit, Panier N-1, Panier | Type d'affichage est Hier · 1 - Axes est Site |

Pièges fréquents (détectés par les contrôles) :

- **Export 1 :** prendre « Libre - Semaine » ou une semaine passée (W43…) au lieu de « Hier ».
- **Export 2 :** colonne Budget vide → pas d'écart au budget. Le lundi, le cumul ne contient que lundi : c'est normal.
- **Export 3 :** laisser Département ou Rayon dans la matrice → le débit est compté plusieurs fois par ticket.
- Les exports 1 et 3 doivent donner le même CA total (même journée) ; sinon la génération est bloquée.

## Le masque PowerPoint

Le code garde la **couverture** du masque (logo, visuel, pied de page) et construit les autres slides
sur la disposition « Titre seul ». Un nouveau masque CFAO fonctionne tant qu'il garde une couverture en slide 1
et une disposition « Titre seul ».

Slides produites : Couverture · L'essentiel · Les secteurs · Où agir · Retours attendus · Comment lire ce flash.
Retours attendus : réponse le lendemain avant 12h (lundi pour un flash du vendredi ou du samedi).

## Réglages (`flash_quotidien/config.py`)

- Seuils d'alerte : CA vs N-1 < -10 %, marge vs N-1 < -10 %, taux < -3 pts, rayon × site avec CA N-1 ≥ 500 000 FCFA.
- Liste des sites : format, zone de pricing (AB/CD), date de conversion.
- Pôles qui répondent (Pricing, Base de données, Promo), heure de réponse, nombre de lignes affichées.
