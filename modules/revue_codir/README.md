# Revue de performance commerciale · mode d'emploi

Outil du lundi : il lit les 3 exports Power BI de la semaine, calcule l'entonnoir Vs N-1, prépare le classeur des chefs de groupe et génère le rapport PowerPoint à la Direction Générale sur le masque de l'entreprise.

## Installation (une seule fois, sur le poste du pôle BDD)

```
pip install -r requirements.txt
streamlit run app.py
```

L'application s'ouvre dans le navigateur (http://localhost:8501). Pour la version PDF du rapport, LibreOffice doit être installé sur le poste ; sinon, ouvrir le PowerPoint puis Fichier › Exporter › PDF.

## Le lundi

| Heure | Qui | Quoi |
|---|---|---|
| 08:00 | Pôle BDD | Tire les 3 exports Power BI de la semaine écoulée (lundi → dimanche) |
| 08:15 | Pôle BDD | Onglet **Accueil** : dépose les 3 exports, vérifie que les 4 contrôles sont verts. Onglet **Rapport PowerPoint** : télécharge le **classeur des chefs de groupe** et le dépose sur Teams |
| avant 10:10 | Chefs de groupe | Remplissent les cellules jaunes « Causes et actions » de leur onglet (3 familles prioritaires + commentaire général) |
| 10:10 | Coordinateur | Onglet **Accueil** : dépose le ou les classeurs complétés. Onglet **Rapport PowerPoint** : télécharge le rapport |
| 10:30 | Coordinateur | Relit la page 2, envoie le rapport au directeur. Le rapport part quoi qu'il arrive : sans saisie, il indique « En attente : présenté en séance par [chef de groupe] » |
| 14:00 | Tous | Point de détail sur le classeur et les plans d'action |

## Les 3 exports attendus

Chaque export est déposé dans **son** emplacement de la barre latérale ; il est contrôlé dès le dépôt (type reconnu, colonnes, semaine).

| Export | Axes en lignes (dans l'ordre) | Mesures à cocher | Filtres |
|---|---|---|---|
| A · Famille × Magasin | Departement, Rayon, Famille, Site nom long | CA, CA N-1, Marge, Marge Hors Promo N-1, Marge Promo N-1, CA Promo, CA Promo N-1, Qté Vente, Qté Vente N-1 | semaine lundi → dimanche, 12 magasins, tous départements, hors e-commerce |
| B · Rayon × Magasin | Département, Site, Rayon | CA, Débit, Débit N-1, Volume, Volume N-1 | 7 jours de la semaine, 12 magasins, départements 01 à 06, HT |
| C · Magasin | Site | Débit, Débit N-1 | 7 jours de la semaine, 12 magasins, HT |

Exporter en .xlsx sans retoucher le fichier : le pied de page « Filtres appliqués » donne la semaine. Dans l'export B, garder les lignes « Total » (le total général donne les tickets du réseau). Repères : A ≈ 1 500 lignes, B ≈ 370 lignes, C = 13 lignes.

## Si un contrôle est rouge

- **CA A ≠ CA B** : un export n'a pas le même périmètre (départements, magasins ou dates). Le retirer et le refaire.
- **Tickets des sites ≠ tickets du réseau** : l'export C n'est pas filtré sur les 12 magasins.
- **Même période** : un export porte sur une autre semaine.
- **Fichier non reconnu** : les colonnes de l'export ont changé dans Power BI ; prévenir le responsable de l'outil.

## Paramètres (en tête de `engine.py`)

- `FORMATS` : code magasin → format (Hyper, Market, Ex-Supeco, Supeco).
- `CONVERSIONS` : dates de conversion des Ex-Supeco, affichées dans le contexte et le lexique.
- `SITES_A_VERIFIER` : magasins signalés « à vérifier » en page 2 (vider la liste quand le point est levé).
- `SEUIL_FAMILLE` (1 M FCFA) et `SEUIL_SIGNIF` (0,5 M FCFA).
- `CHEFS` : chef de groupe responsable de chaque secteur.

## Contenu

- `app.py` : application Streamlit (Accueil et lexique, Synthèse exécutive, Drill-down secteurs, Rapport PowerPoint).
- `engine.py` : lecture des exports, contrôles et tous les calculs (source unique des chiffres).
- `ppt_report.py` : rapport de 5 pages + annexe sur le masque `assets/masque_cfao_consumer.pptx`.
- `excel_cdg.py` : classeur des chefs de groupe (valeurs uniquement, aucune formule) et relecture des « Causes et actions ».
