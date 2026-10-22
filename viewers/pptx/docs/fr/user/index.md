---
title: Aperçu du guide utilisateur
description: Une introduction à l'éditeur pptx-viewer - les parties de l'écran, comment ouvrir un fichier et où trouver chaque fonctionnalité.
---

# Guide utilisateur

Ce guide s'adresse aux personnes qui **utilisent** l'application pptx-viewer pour travailler avec des présentations PowerPoint (`.pptx`) dans le navigateur. Il couvre l'affichage, l'édition, la présentation, l'exportation et la collaboration. Aucun code n'est nécessaire.

Si vous êtes développeur et intégrez la visionneuse dans votre propre application, consultez plutôt la [documentation développeur de la visionneuse React](/react/).

::: info Les fonctionnalités dépendent de la configuration de l'application
pptx-viewer est un composant configurable : les boutons et panneaux exacts que vous voyez dépendent donc de la façon dont le développeur de votre application a activé les fonctionnalités. L'édition n'est disponible que si l'application l'active (un badge **Lecture seule** apparaît dans la barre d'outils dans le cas contraire), et les commandes de collaboration n'apparaissent que lorsque le partage est activé. Lorsqu'une fonctionnalité est optionnelle, ce guide l'indique.
:::

## Ce qu'est l'éditeur

pptx-viewer ouvre les fichiers PowerPoint directement dans votre navigateur et affiche chaque diapositive en vrai HTML et SVG, et non en image plate. Le texte reste net à tout niveau de zoom, peut être sélectionné et copié, et fonctionne avec les lecteurs d'écran. En mode édition, vous pouvez modifier les diapositives et enregistrer le résultat dans un fichier `.pptx` standard.

![L'éditeur avec une présentation ouverte](/docs-shots/editor.jpg)

## Ouvrir un fichier

1. Faites glisser un fichier `.pptx` sur la zone de dépôt intitulée **« Drop a .pptx file here or click to browse »**, ou cliquez sur la zone pour ouvrir le sélecteur de fichiers de votre système.
2. Patientez un instant pendant l'analyse du fichier. La première diapositive apparaît sur le canevas, avec les vignettes à gauche.

Vous pouvez aussi cliquer sur **« or create a New Presentation »** pour partir d'une présentation vierge d'une seule diapositive.

::: tip Votre fichier reste sur votre machine
Le fichier est traité entièrement dans le navigateur. Rien n'est envoyé à un serveur : l'ouverture, l'édition et l'enregistrement se font tous en local.
:::

::: warning Fichiers pris en charge
L'application ouvre les fichiers `.pptx` et leurs variantes (`.ppsx`, `.pptm`, `.potx`), ainsi que le **format binaire historique `.ppt`** de PowerPoint 97-2003, qui est converti au chargement. L'enregistrement se fait par défaut en `.pptx`, quel que soit le fichier ouvert (une présentation ouverte sous `report.ppt` est proposée sous `report.pptx`), mais Enregistrer sous propose aussi **Présentation PowerPoint 97-2003 (`.ppt`)** comme choix de format explicite, à côté de `.ppsx` et `.pptm` ; voir [Limitations](/fr/guide/limitations) pour les écarts de fidélité restants sur ce format (actions de clic sur les liens hypertexte et éléments sans équivalent binaire).

Les fichiers `.pptx` protégés par mot de passe s'ouvrent dès que vous fournissez le mot de passe : une boîte de dialogue intitulée **Fichier chiffré** le demande. Les fichiers **`.ppt`** protégés par mot de passe ne peuvent pas être ouverts du tout, car l'ancien schéma de chiffrement n'est pas déchiffré.
:::

Si l'application trouve un travail non enregistré d'une visite précédente, une bannière intitulée **Modifications non enregistrées récupérées** apparaît avec les boutons **Restaurer** et **Ignorer**. Voir [Sauvegarde automatique et récupération](/fr/user/editing#sauvegarde-automatique-et-recuperation).

## Visite de l'écran

![Disposition complète de l'éditeur](/user-guide/overview-full-layout.jpg)

De haut en bas :

### Barre de titre

La bande supérieure affiche le nom du fichier et son état d'enregistrement (par exemple « Enregistré sur ce PC » ou « Modifications non enregistrées »). Lorsque l'édition est activée, elle contient aussi :

- Le commutateur **Sauvegarde automatique** (Activée/Désactivée).
- Les boutons rapides **Enregistrer**, **Annuler** et **Rétablir**.
- Une zone de recherche, **« Dis-moi ce que tu veux faire »**. Tapez le nom d'une commande (par exemple « tableau ») pour obtenir une liste cliquable d'actions correspondantes, ou appuyez sur **Entrée** pour rechercher le texte dans vos diapositives.

### Ruban (barre d'outils)

Un ruban à onglets sous la barre de titre. Les onglets, dans l'ordre :

**Fichier, Accueil, Insertion, Dessin, Création, Transitions, Animations, Diaporama, Enregistrer, Révision, Affichage, Aide**

![Gros plan de la barre d'outils](/user-guide/overview-toolbar.jpg)

- **Fichier** ouvre un écran Backstage plein écran avec les pages Ouvrir, Enregistrer, Enregistrer sous, Imprimer, Partager et Exporter. Voir [Exportation](/fr/user/exporting).
- **Accueil** regroupe les outils courants : presse-papiers, nouvelles diapositives et dispositions, polices, mise en forme des paragraphes, rechercher et remplacer, dessin et commandes d'organisation. Voir [Édition](/fr/user/editing).
- **Insertion** ajoute de nouveaux éléments : zones de texte, formes, images, médias, tableaux, graphiques, SmartArt, équations, etc.
- **Diaporama** lance une présentation. Voir [Présentation](/fr/user/presenting).

Le ruban n'apparaît qu'en mode édition. À l'extrémité droite de la rangée d'onglets se trouvent les boutons **Enregistrer** et **Partager** ; la rangée du dessus contient des commutateurs pour le panneau des diapositives, les commentaires et l'inspecteur, le bouton **Présenter**, un engrenage Paramètres et un menu de débordement **« ... »** qui reprend les actions d'exportation et les utilitaires.

### Panneau des diapositives (gauche)

Une bande défilante de vignettes de diapositives.

![Panneau des diapositives](/user-guide/overview-slides-panel.jpg)

- Cliquez sur une vignette pour accéder à cette diapositive.
- Faites glisser les vignettes pour les réorganiser (mode édition).
- Les diapositives masquées apparaissent estompées, avec des hachures diagonales et une icône d'œil barré.
- Le bouton **Ajouter une diapositive** en bas insère une nouvelle diapositive.
- Les diapositives peuvent être regroupées en **sections** nommées, avec leurs propres en-têtes.

### Canevas de diapositive (centre)

La grande zone où s'affiche la diapositive active. Vous y zoomez et vous y déplacez, et, en mode édition, vous sélectionnez, déplacez et modifiez directement les éléments. Les règles, une grille et des repères déplaçables peuvent être activés depuis l'onglet **Affichage**.

### Inspecteur (droite)

Un panneau latéral contextuel à trois onglets :

- **Éléments** - une liste en calques de tous les éléments de la diapositive, dans l'ordre d'empilement. Cliquez sur une ligne pour sélectionner cet élément sur le canevas.
- **Propriétés** - les paramètres de la sélection en cours. Sans sélection, il affiche les paramètres de la présentation, du thème, de la taille des diapositives, de la transition et de l'arrière-plan ; avec un élément sélectionné, il affiche la position et la taille ainsi que des panneaux propres au type d'élément (texte, remplissage et contour, tableau, données du graphique, image, média, etc.) et un panneau **Animation**.
- **Commentaires** - les fils de commentaires attachés à la diapositive en cours, avec des commandes pour répondre et résoudre.

Affichez ou masquez l'inspecteur avec le bouton de panneau de la rangée supérieure de la barre d'outils.

### Panneau de notes (sous le canevas)

Un éditeur pour les notes du présentateur de la diapositive en cours. Affichez-le ou masquez-le avec le bouton **Notes** de la barre d'état.

### Barre d'état (bas)

De gauche à droite :

- Le compteur **Diapositive x sur y**.
- L'état d'enregistrement : **Modifications non enregistrées**, **Tout est enregistré**, **Enregistrement...** ou **Enregistré [heure]**.
- Le commutateur **Notes**.
- Trois boutons d'affichage : **Affichage normal**, **Trieuse de diapositives**, **Diaporama**.
- L'indicateur de connexion lorsqu'une session de collaboration est active.
- Les commandes de zoom : **Zoom arrière**, un bouton de pourcentage qui sert aussi de **Ajuster à la fenêtre**, et **Zoom avant**.

## Ordinateur et mobile

::: warning Conçu pour l'ordinateur
Le ruban, l'inspecteur et les boîtes de dialogue sont conçus pour les écrans de taille d'ordinateur. Sur un téléphone ou une petite tablette, l'application passe à une barre d'outils compacte et à des feuilles inférieures, et prend en charge les gestes tactiles (glisser, pincer pour zoomer), mais une édition sérieuse est plus facile sur un grand écran.
:::

<div style="display: flex; gap: 1rem; align-items: flex-start;">
  <figure style="flex: 1;">
    <img src="/user-guide/overview-mobile-layout.jpg" alt="Disposition mobile de la visionneuse" style="max-width: 280px;" />
    <figcaption>Visionneuse mobile</figcaption>
  </figure>
  <figure style="flex: 1;">
    <img src="/user-guide/overview-mobile-inspector.jpg" alt="Feuille inférieure de l'inspecteur mobile" style="max-width: 280px;" />
    <figcaption>Inspecteur mobile (feuille inférieure)</figcaption>
  </figure>
</div>

## Où aller ensuite

- [Affichage des présentations](/fr/user/viewing) - naviguer dans les diapositives, zoom, diapositives masquées et lecture des notes.
- [Édition des diapositives](/fr/user/editing) - sélectionner et mettre en forme des éléments, insérer du contenu, gérer les diapositives, annuler et sauvegarde automatique.
- [Présentation](/fr/user/presenting) - lancer un diaporama, vue du présentateur et annotations en direct.
- [Exportation](/fr/user/exporting) - enregistrer en `.pptx`, exporter en PDF, images, GIF et vidéo, et imprimer.
- [Collaboration](/fr/user/collaboration) - coéditer en temps réel et diffuser un diaporama.
- [Raccourcis clavier](/fr/user/shortcuts) - la référence complète des raccourcis.
