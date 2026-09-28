# Portfolio BTS SIO SISR – Kylian Grafeille Clément

Site statique (HTML / CSS / JavaScript, sans dépendance) servant de support à la présentation du BTS SIO option SISR.

## Lancer le site

Ouvrir `index.html` dans un navigateur, ou le publier avec **GitHub Pages** :
*Settings → Pages → Source : branche `main`, dossier `/ (root)`*.

## Structure

```
index.html          contenu du portfolio (toutes les sections)
css/style.css       mise en forme (thème sombre « informatique »)
js/main.js          menu, animations, zoom des captures, emplacements d'images
assets/icons/       icônes des outils et des réseaux (SVG)
assets/img/         photo, logo et captures d'écran
```

## Ce qu'il reste à compléter

Les zones surlignées en jaune **« À compléter »** dans `index.html` sont à remplir
(organisation de SIKA, statut / date du CCNA). Rien n'a été inventé.

## Images à ajouter

Tant qu'une image est absente, le site affiche un cadre indiquant le chemin attendu.
Il suffit de déposer le fichier avec le bon nom :

| Fichier | Contenu |
|---|---|
| `assets/img/photo.jpg` | Photo de profil (carrée, ~600×600 px) |
| `assets/img/logo-sika.png` | Logo SIKA France |
| `assets/img/veille/feedly.png` | Capture de l'espace Feedly |

## Ajouter un lien dans « Connecte »

Exemple : ajouter Instagram.

**1. Récupérer l'icône de l'application**

- Aller sur [simpleicons.org](https://simpleicons.org), chercher l'application (ex. « Instagram »).
- Cliquer sur l'icône puis **Download SVG**.
- Renommer le fichier en minuscules, sans espace (ex. `instagram.svg`) et le placer dans `assets/icons/`.
- Noter la couleur de la marque affichée sur le site (ex. `#FF0069`).

> Si l'application n'y est pas (c'est le cas de LinkedIn), chercher sur
> [fontawesome.com/search?ic=brands](https://fontawesome.com/search?ic=brands) → télécharger le SVG.
> Peu importe la couleur de l'icône : le site l'affiche automatiquement en blanc.

**2. Ajouter le bloc dans `index.html`**

Chercher `<div class="socials">` (section `CONNECTE`) et copier ce bloc juste avant `</div>` :

```html
<a class="social" href="https://www.instagram.com/TON_PSEUDO/" target="_blank" rel="noopener" style="--brand:#FF0069">
  <img src="assets/icons/instagram.svg" alt="">
  <span>Instagram</span>
</a>
```

Il y a 4 choses à modifier :

| Élément | À remplacer par |
|---|---|
| `href="..."` | le lien vers ton profil |
| `--brand:#FF0069` | la couleur de la marque (fond de l'icône) |
| `src="assets/icons/instagram.svg"` | le nom du fichier de l'icône |
| `<span>Instagram</span>` | le nom affiché sous l'icône |

**3. Vérifier** : ouvrir `index.html` dans le navigateur (ou recharger avec `Ctrl + F5`).

Pour un e-mail, utiliser `href="mailto:adresse@exemple.fr"` (sans `target="_blank"`).

La même méthode fonctionne pour ajouter un outil dans la section **Outils** : copier un bloc `<div class="item">`.

## Crédits

Icônes : [Simple Icons](https://simpleicons.org) (CC0) et [Font Awesome Free](https://fontawesome.com) (CC BY 4.0).

## Présentation PowerPoint

`presentation/Presentation-BTS-SIO-SISR-Kylian-Grafeille.pptx` reprend tout le contenu du portfolio (15 diapos, même thème).
Elle se modifie directement dans PowerPoint : remplacer le cercle « Photo de profil » et le cadre « Logo SIKA France »
par les images, et remplir les zones jaunes.

`presentation/generer-presentation.js` est le script qui a généré le fichier (facultatif) :
`npm install pptxgenjs sharp react react-dom react-icons` puis `node presentation/generer-presentation.js`.
Attention : relancer le script écrase les modifications faites à la main dans le .pptx.
