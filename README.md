# SmartBuyer Hub · NovaRetail

Plateforme Streamlit de l'équipe Achats : chaque module lit des exports Power BI et produit une analyse ou un livrable (Excel, PowerPoint, HTML). Aucune donnée n'est stockée dans ce dépôt.

## Lancer en local

```
pip install -r requirements.txt
streamlit run app.py
```

## Organisation du dépôt

```
app.py                 Accueil du hub : registre des modules (META) et cartes par rubrique
pages/                 Un fichier par module actif : NN_<emoji>_Nom_Du_Module.py
modules/               Moteurs des modules lourds, un dossier par module (revue_codir/…)
.streamlit/            Configuration Streamlit (les secrets restent dans Streamlit Cloud)
requirements.txt       Dépendances
```

Les modules retirés sont conservés sur la branche `archive` (récupérables à tout moment) : Scoring ABC, Ventes PBI, Détention Top CA, Performance Promo, OOS, Tasks Tracker, Perf Hebdo, Rentabilité, Fidélité Cagnotte, Marge, Perf Review, Bascule XD, COPIL Hebdo, Reporting Ventes, Reporting Promo, COPIL Journalier PGC, Cockpit Promo, Synthèse PGC.

## Modules actifs

| N° | Module | Rubrique |
|---|---|---|
| 26 | Revue CODIR | Pilotage |
| 24 | PGC Radar | Pilotage |
| 18 | Reporting Vente CA | Ventes et marge |
| 17 | Reporting Sous-Familles | Ventes et marge |
| 22 | Suivi Rayon vs Pays | Ventes et marge |
| 06 | Marges Négatives | Ventes et marge |
| 05 | Suivi Implantation | Assortiment et fournisseurs |
| 07 | OTIF | Assortiment et fournisseurs |

## Conventions

- **Nom de fichier** : `NN_<emoji>_Nom_Du_Module.py`, sans espace ; l'emoji est celui de `st.set_page_config(page_icon=...)`.
- **Nouveau module** : ajouter le fichier dans `pages/` et une ligne dans `META` (`app.py`). Un module absent de `META` s'affiche quand même dans « Autres modules ».
- **Code partagé** : créer un dossier `core/` le jour où deux modules partagent du code (`from core import ...`).
- **Gros module** : la page reste dans `pages/`, les calculs vont dans `modules/<nom>/`.
- **Archiver un module** : le supprimer de `pages/` et de `META` ; il reste sur la branche `archive` et dans l'historique.
- **Libellés** en français, code interne en anglais.
- **Exports Excel** : valeurs uniquement, aucune formule.
- **Données et documents internes** (exports, masques PowerPoint) : jamais dans le dépôt, toujours déposés dans l'application.
