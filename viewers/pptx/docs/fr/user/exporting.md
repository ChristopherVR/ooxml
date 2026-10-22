---
title: Exportation
description: Enregistrez vos modifications en .pptx, exportez en PDF, PNG, GIF et vidéo, préparez un package de partage et imprimez des documents et des pages de notes.
---

# Exportation

Tout ce que vous exportez est généré dans votre navigateur et téléchargé sur votre appareil ; rien n'est envoyé ailleurs. Les actions d'exportation et d'enregistrement se trouvent dans l'onglet **Fichier** (le Backstage), et les plus courantes sont reprises dans le menu de débordement **« ... »** à l'extrémité droite de la barre d'outils.

![Onglet Fichier affichant les options d'exportation](/user-guide/exporting-file-menu.jpg)

## Enregistrer au format PowerPoint

1. Cliquez sur l'onglet **Fichier**.
2. Cliquez sur **Enregistrer** pour télécharger la présentation en fichier `.pptx`, ou ouvrez **Enregistrer sous** pour tous les formats :

| Option Enregistrer sous                 | Ce que vous obtenez                                                                                          |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| **Présentation PowerPoint**             | Un fichier `.pptx` modifiable avec toutes vos modifications.                                                 |
| **Diaporama PowerPoint**                | Un fichier `.ppsx` qui s'ouvre directement en diaporama.                                                     |
| **Présentation activée par des macros** | Un fichier `.pptm` qui préserve les macros VBA. Affiché uniquement si le fichier ouvert contient des macros. |
| **Package de partage**                  | Un lot `.zip` contenant le `.pptx` et un fichier readme.                                                     |

Le bouton **Enregistrer** de la barre de titre fait la même chose que **Fichier > Enregistrer**. Les fichiers enregistrés conservent vos modifications lors de l'aller-retour : éléments ajoutés et modifiés, ordre des diapositives, notes, sections, etc.

::: warning Il n'y a pas de Ctrl+S
Comme l'application s'exécute dans un navigateur, **Ctrl+S** déclenche la boîte de dialogue d'enregistrement de page du navigateur, et non un enregistrement de la présentation. Utilisez plutôt le bouton **Enregistrer** ou l'onglet Fichier. Les fichiers sont toujours enregistrés sous forme de téléchargements ; consultez le dossier de téléchargement de votre navigateur.
:::

## Exporter des images, des PDF, des GIF et des vidéos

Ouvrez **Fichier > Exporter** :

| Option                               | Ce que vous obtenez                                                                | Portée               |
| ------------------------------------ | ---------------------------------------------------------------------------------- | -------------------- |
| **Créer un PDF**                     | Un PDF d'une page par diapositive, capturé en haute résolution.                    | Présentation entière |
| **Exporter la diapositive en cours** | Une image PNG de la diapositive en cours.                                          | Diapositive en cours |
| **Créer une vidéo**                  | Une vidéo `.webm` qui parcourt la présentation, quelques secondes par diapositive. | Présentation entière |
| **Créer un GIF animé**               | Un GIF compact en boucle de la présentation.                                       | Présentation entière |
| **Exporter en JSON**                 | Un document JSON portable qui se réimporte avec une fidélité complète.             | Présentation entière |
| **Copier en tant qu'image**          | Copie la diapositive en cours dans le presse-papiers sous forme d'image.           | Diapositive en cours |

Pour les formats de présentation entière, une boîte de dialogue de progression montre chaque diapositive en cours de capture, avec un bouton **Annuler** si vous changez d'avis. Le fichier terminé est téléchargé automatiquement.

## Impression

1. Ouvrez **Fichier > Imprimer** et cliquez sur **Imprimer la présentation**. Une boîte de dialogue d'impression s'ouvre avec un aperçu en direct.
2. Choisissez ce qu'il faut imprimer sous **Imprimer** : **Diapositives pleine page**, **Documents** (avec un réglage du nombre de diapositives par page), **Pages de notes** (chaque diapositive au-dessus de ses notes du présentateur) ou **Plan**.
3. Définissez l'**Orientation**, le **Mode couleur** (Couleur, Niveaux de gris, Noir et blanc), s'il faut **Encadrer les diapositives**, et la **Plage de diapositives** (Toutes les diapositives, Diapositive en cours, ou une plage De/À personnalisée).
4. Cliquez sur **Imprimer**. La fenêtre d'impression de votre navigateur s'ouvre ; vous y choisissez l'imprimante ou l'enregistrement en PDF.

::: tip Documents de notes
Imprimer des **Pages de notes** est le moyen d'obtenir un document qui associe chaque diapositive à ses notes du présentateur.
:::

## Protéger et inspecter le fichier

La page **Fichier > Informations** propose :

- **Protéger la présentation** - définir un mot de passe pour que le fichier enregistré soit chiffré.
- **Inspecter la présentation** - examiner les propriétés du document.
- **Incorporer les polices** - garder une typographie cohérente lorsque le fichier passe d'un appareil à un autre.
- **Signatures numériques** - afficher les signatures attachées à la présentation.

## Notes sur la fidélité et la taille

::: tip Les exports raster correspondent de près à l'écran
Les exports PNG, PDF, GIF et vidéo capturent la diapositive affichée à l'écran sous forme de document SVG autonome et laissent le moteur de rendu de votre navigateur le dessiner ; des effets comme les arrière-plans flous/transparents, les couleurs personnalisées et les biseaux 3D sont donc rendus au lieu d'être approximés. `html2canvas` n'est utilisé qu'en repli, pour la rare ressource (une image inter-origine sans identifiants, ou une police qui ne se charge pas) qui ne peut pas être capturée ainsi. Voir [Limitations](/fr/guide/limitations) pour le petit ensemble d'approximations CSS restantes.
:::

::: info Très grandes diapositives
Un export dont la résolution dépasserait la taille maximale de canevas de votre navigateur est automatiquement découpé en tuiles puis réassemblé dans le fichier final ; une diapositive extrêmement grande exportée à une échelle élevée produit donc tout de même un PNG, un PDF, un GIF ou une vidéo complet et de la bonne taille, au lieu d'être plafonné ou tronqué. Aucun réglage à modifier : cela se fait de façon transparente.
:::

## Suite

- Coéditer avec d'autres : [Collaboration](/fr/user/collaboration)
- Détails d'export pour les développeurs : [/react/export](/react/export)
