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
core/                  Code partagé entre modules (utils_promo.py…)
modules/               Moteurs des modules lourds, un dossier par module (revue_codir/…)
.streamlit/            Configuration Streamlit (les secrets restent dans Streamlit Cloud)
requirements.txt       Dépendances
```

Les anciens modules sont conservés sur la branche `archive` (récupérables à tout moment).

## Modules actifs

| N° | Module | Rubrique |
|---|---|---|
| 05 | Suivi Implantation | Assortiment & stock |
| 06 | Marges Négatives | Performance commerciale |
| 07 | OTIF | Fournisseurs & organisation |
| 12 | Bascule XD | Fournisseurs & organisation |
| 15 | COPIL Hebdo | Reporting & COPIL |
| 16 | Reporting Ventes | Reporting & COPIL |
| 17 | Reporting Sous-Familles | Reporting & COPIL |
| 18 | Reporting Vente CA | Reporting & COPIL |
| 19 | Reporting Promo | Reporting & COPIL |
| 21 | COPIL Journalier PGC | Reporting & COPIL |
| 22 | Suivi Rayon vs Pays | Performance commerciale |
| 23 | Cockpit Promo | Performance commerciale |
| 24 | PGC Radar | Performance commerciale |
| 25 | Synthèse PGC | Reporting & COPIL |
| 26 | Revue CODIR | Reporting & COPIL |

## Conventions

- **Nom de fichier** : `NN_<emoji>_Nom_Du_Module.py`, sans espace ; l'emoji est celui de `st.set_page_config(page_icon=...)`.
- **Nouveau module** : ajouter le fichier dans `pages/` et une ligne dans `META` (`app.py`). Un module absent de `META` s'affiche quand même dans « Autres modules ».
- **Code partagé** : dans `core/`, importé avec `from core import ...`.
- **Gros module** : la page reste dans `pages/`, les calculs vont dans `modules/<nom>/`.
- **Archiver un module** : le supprimer de `pages/` et de `META` ; il reste sur la branche `archive` et dans l'historique.
- **Libellés** en français, code interne en anglais.
- **Exports Excel** : valeurs uniquement, aucune formule.
- **Données et documents internes** (exports, masques PowerPoint) : jamais dans le dépôt, toujours déposés dans l'application.
