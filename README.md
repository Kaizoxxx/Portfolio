# Portfolio BTS SIO SISR – Kylian Grafeille Clément

Site statique (HTML / CSS / JavaScript, sans dépendance) servant de support à la présentation du BTS SIO option SISR.

## Lancer le site

Ouvrir `index.html` dans un navigateur, ou le publier avec **GitHub Pages** :
*Settings → Pages → Source : branche `main`, dossier `/ (root)`*.

## Structure

```
index.html          contenu du portfolio (toutes les sections)
css/style.css       mise en forme (thème clair / sombre automatique)
js/main.js          menu, onglets, zoom des captures, emplacements d'images
assets/img/         photos, logos et captures d'écran
```

## Ce qu'il reste à compléter

Les zones surlignées en jaune **« À compléter »** dans `index.html` sont à remplir
(missions en entreprise, projets de formation, projet de 2e année). Rien n'a été inventé.

## Images à ajouter

Tant qu'une image est absente, le site affiche un cadre indiquant le chemin attendu.
Il suffit de déposer le fichier avec le bon nom :

| Fichier | Contenu |
|---|---|
| `assets/img/photo.jpg` | Photo de profil (carrée, ~600×600 px) |
| `assets/img/logo-sika.png` | Logo SIKA France |
| `assets/img/missions/mission1-1.png`, `mission1-2.png` … `mission3-2.png` | Captures des missions en entreprise |
| `assets/img/projets/bac-1.png` | Capture projet Bac Pro SN |
| `assets/img/projets/bts1-1.png` | Capture projet BTS 1re année |
| `assets/img/projets/bts2-1.png` | Capture projet BTS 2e année |
| `assets/img/projet-e2/schema-reseau.png` | Schéma réseau du projet de 2e année |
| `assets/img/projet-e2/capture-1.png` … `capture-3.png` | Captures du projet de 2e année |
| `assets/img/veille/feedly.png` | Capture de l'espace Feedly |

Pour ajouter une mission ou un projet : dupliquer le bloc `<article class="mission">` (ou `<article class="project">`) dans `index.html`.
