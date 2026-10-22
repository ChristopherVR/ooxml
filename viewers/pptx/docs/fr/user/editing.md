---
title: Édition des diapositives
description: Sélectionnez, déplacez, redimensionnez et mettez en forme des éléments ; insérez du texte, des formes, des images, des tableaux et des graphiques ; gérez les diapositives et les sections ; annulez ; et récupérez un travail non enregistré.
---

# Édition des diapositives

Cette page couvre la modification du contenu d'une présentation : travailler avec les éléments, la mise en forme, l'organisation, la gestion des diapositives et la protection de votre travail.

::: warning L'édition doit être activée
L'édition n'est disponible que si l'application est configurée pour cela. Si la barre d'outils affiche un badge **Lecture seule** et que vous ne pouvez pas sélectionner ou déplacer d'éléments, l'application est en mode consultation uniquement. Développeurs : voir la documentation des [Props des composants](/react/props).
:::

## Sélectionner des éléments

- **Cliquez** sur un élément du canevas pour le sélectionner. Des poignées apparaissent autour de lui.
- **Maj-clic** sur d'autres éléments pour les ajouter à la sélection.
- **Faites glisser sur le canevas vide** pour tracer un cadre de sélection et sélectionner tout ce qu'il contient.
- Appuyez sur **Ctrl+A** (Cmd sur macOS) pour sélectionner tous les éléments de la diapositive.
- Appuyez sur **Échap** pour désélectionner.

Vous pouvez aussi sélectionner depuis l'onglet **Éléments** de l'inspecteur, qui liste tous les éléments de la diapositive dans l'ordre d'empilement sous le titre **Ordre des calques**. C'est pratique lorsque des éléments se chevauchent.

![Élément sélectionné avec ses poignées de redimensionnement visibles](/user-guide/editing-element-selected.jpg)

## Déplacer, redimensionner et faire pivoter

Avec un élément sélectionné :

- **Déplacer** - faites-le glisser. Des repères d'alignement et des lignes d'aimantation apparaissent lorsqu'il s'aligne sur la grille ou sur d'autres formes.
- **Décaler légèrement** - appuyez sur les **touches fléchées** pour de petits pas, **Maj + touches fléchées** pour de plus grands pas.
- **Redimensionner** - faites glisser n'importe quelle poignée de coin ou de bord.
- **Faire pivoter** - faites glisser la poignée de rotation au-dessus de l'élément. Maintenez **Maj** pendant la rotation pour aimanter par pas de 15 degrés.
- **Ajuster la forme** - certaines formes affichent une poignée d'ajustement supplémentaire qui modifie leur géométrie (par exemple la taille de la pointe d'une flèche).

![Déplacement et redimensionnement d'éléments](/user-guide/editing-drag-resize.gif)

Pour un placement exact, ouvrez l'onglet **Propriétés** de l'inspecteur. La carte **Élément** comporte les champs **X**, **Y**, **L** et **H**, ainsi qu'un commutateur de verrouillage qui protège l'élément contre les glissements accidentels.

::: tip Règles, grille et repères
Activez **Règles**, **Grille**, **Repères** et **Aligner sur la grille** depuis le groupe Afficher de l'onglet **Affichage**. Faites glisser un repère pour le repositionner ; double-cliquez sur un repère pour le supprimer.
:::

## Modifier le texte

1. **Double-cliquez** sur une zone de texte, une forme ou une cellule de tableau. Un curseur apparaît et vous pouvez saisir du texte.
2. Appuyez sur **Ctrl+Entrée** pour valider la modification, ou sur **Échap** pour l'annuler.

![Édition de texte en ligne avec le curseur actif](/user-guide/editing-inline-text.jpg)

Mettez le texte en forme depuis l'onglet **Accueil** :

- **Groupe Police** - listes déroulantes de la famille et de la taille de police ; **Gras**, **Italique**, **Souligné**, **Barré** ; **Ombre du texte** ; **Augmenter / Réduire la taille de police** ; **Effacer la mise en forme** ; **Espacement des caractères** ; **Modifier la casse** ; **Couleur de police** et **Couleur de surbrillance du texte** (chacune avec des échantillons prédéfinis et un sélecteur **Couleur personnalisée...**).

  La liste des polices est regroupée comme le fait PowerPoint : les deux
  **polices de thème** de la présentation d'abord (marquées _Titres_ et _Corps_), puis
  les polices **incorporées dans cette présentation**, puis celles que vous avez
  [ajoutées pendant cette session](#ajouter-une-police-que-la-visionneuse-ne-possede-pas), puis la liste
  alphabétique complète. Si rien ne la remplace sur l'élément sélectionné, la zone
  affiche la police que la présentation utiliserait réellement, qui, dans un espace
  réservé de titre, est la police de titre du thème plutôt qu'une valeur générique par défaut.

- **Groupe Paragraphe** - listes à puces et numérotées ; commandes de retrait ; **Aligner à gauche / au centre / à droite / Justifier** ; **Interligne** ; **Orientation du texte** (horizontale, pivotée, empilée) ; **Colonnes**.

Les mêmes paramètres de texte sont disponibles dans le panneau **Texte** de l'inspecteur lorsqu'un élément texte est sélectionné.

## Insérer des éléments

Utilisez l'onglet **Insertion** pour ajouter du contenu à la diapositive en cours :

![Onglet Insertion avec les options d'éléments](/user-guide/editing-insert-tab.jpg)

| Bouton                      | Ce qu'il insère                                                                                                                                           |
| --------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Zone de texte**           | Une zone de texte vide prête à la saisie.                                                                                                                 |
| **Forme**                   | Une forme prédéfinie ; choisissez le type de forme dans la liste déroulante à côté du bouton (rectangles, flèches, légendes, étoiles, connecteurs, etc.). |
| **Image**                   | Une image provenant d'un fichier de votre appareil.                                                                                                       |
| **Média**                   | Un fichier vidéo ou audio.                                                                                                                                |
| **Tableau**                 | Un nouveau tableau sur la diapositive.                                                                                                                    |
| **Graphique**               | Un graphique ; choisissez **Barres**, **Courbes**, **Secteurs**, **Anneau**, **Aires** ou **Nuage de points** dans la liste déroulante à côté du bouton.  |
| **SmartArt**                | Un diagramme SmartArt choisi dans une galerie de dispositions.                                                                                            |
| **Équation**                | Une équation mathématique construite dans l'éditeur d'équations.                                                                                          |
| **Action**                  | Un bouton d'action (flèches de navigation, accueil et préréglages similaires).                                                                            |
| **Champ**                   | Un champ dynamique : **Numéro de diapositive**, **Date/Heure**, **En-tête** ou **Pied de page**.                                                          |
| **En-tête et pied de page** | Ouvre la boîte de dialogue d'en-tête et de pied de page de la présentation.                                                                               |

Les nouveaux éléments arrivent sur la diapositive sélectionnés et prêts à être déplacés, redimensionnés et mis en forme.

![Formes sur une diapositive](/docs-shots/shapes-slide.jpg)

Il existe aussi un onglet **Dessin** avec des outils à main levée : **Sélectionner**, **Stylo**, **Surligneur**, **Gomme** et **Forme libre**, ainsi qu'un sélecteur de couleur de stylo et un curseur d'épaisseur de trait. L'encre que vous dessinez devient des éléments de la diapositive.

## Mettre en forme avec l'inspecteur

L'onglet **Propriétés** de l'inspecteur change selon la sélection :

| Lorsque vous sélectionnez... | Vous pouvez modifier...                                                                                                                             |
| ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| Rien (la diapositive)        | Informations de la présentation, thème, taille des diapositives, transition et arrière-plan (couleur ou image).                                     |
| Une forme ou un connecteur   | Position, taille, rotation, remplissage (uni, dégradé, motif, image) et contour.                                                                    |
| Un élément texte             | Police, taille, couleur, alignement et espacement.                                                                                                  |
| Une image                    | Rognage et retouches de l'image.                                                                                                                    |
| Un tableau                   | Styles de tableau prédéfinis, largeurs de colonnes, hauteurs de lignes et mise en forme des cellules.                                               |
| Un graphique                 | La grille de données des séries, le type de graphique, les axes, les étiquettes de données, les marqueurs, les courbes de tendance et les couleurs. |
| SmartArt                     | Texte des nœuds et disposition.                                                                                                                     |
| Média                        | Paramètres de lecture.                                                                                                                              |
| Tout élément                 | **Animations** d'entrée, d'accentuation et de sortie, avec direction, minutage et déclencheur.                                                      |

![Panneau d'inspection affichant les propriétés d'un élément](/user-guide/editing-inspector-panel.jpg)

### Modifier les graphiques sur le canevas

Les graphiques sont modifiables directement, sans quitter la diapositive :

- **Survolez** une marque pour voir sa valeur dans une info-bulle.
- **Cliquez** sur une barre, un point ou un secteur pour le sélectionner ; la cellule correspondante est mise en évidence dans la grille de données de l'inspecteur.
- **Faites glisser** une barre ou un point vers le haut ou le bas pour modifier sa valeur (graphiques en barres, courbes, aires, nuages de points et bulles). Une pastille affiche la valeur pendant le glissement ; relâchez pour valider ou appuyez sur **Échap** pour annuler.
- **Double-cliquez sur le titre du graphique** pour le renommer sur place.

Les marques des graphiques en secteurs, en anneau, en radar et empilés/empilés à 100 % se sélectionnent au clic ; modifiez leurs valeurs dans la grille de données. Chaque modification est une seule étape d'annulation. Tout cela fonctionne de la même façon dans tous les bindings (React, Vue, Angular, Svelte et vanilla).

### Modifier les tableaux

Faites un clic droit sur un tableau pour accéder aux commandes de lignes et de colonnes : **Insérer une ligne au-dessus / en dessous**, **Insérer une colonne à gauche / à droite**, **Supprimer la ligne**, **Supprimer la colonne**, **Fusionner les cellules** et **Fractionner la cellule**. Double-cliquez sur une cellule pour modifier son texte.

![Un tableau sur une diapositive](/docs-shots/table-slide.jpg)

## Organiser les éléments

![Commandes d'organisation dans le ruban](/user-guide/editing-arrange-tab.jpg)

- **Ordre d'empilement** - faites un clic droit sur un élément et choisissez **Avancer**, **Reculer**, **Premier plan** ou **Arrière-plan**. Les mêmes commandes se trouvent dans le groupe Dessin de l'onglet Accueil sous **Organiser**, et sous forme de boutons de calques dans l'inspecteur.
- **Aligner et répartir** - avec au moins deux éléments sélectionnés, utilisez les boutons d'alignement de l'onglet Accueil (gauche, centre, droite, haut, milieu, bas) et **Répartir** horizontalement ou verticalement.
- **Grouper / Dissocier** - faites un clic droit sur une sélection multiple et choisissez **Grouper** pour que les éléments se déplacent et se redimensionnent ensemble ; **Dissocier** les sépare de nouveau.
- **Retourner** - utilisez **Retourn. H** / **Retourn. V** pour mettre la sélection en miroir.

## Gérer les diapositives

- **Ajouter une diapositive** - cliquez sur **Ajouter une diapositive** en bas du panneau des diapositives, ou utilisez **Accueil > Nouvelle diapositive**. La flèche à côté de Nouvelle diapositive ouvre une galerie des dispositions appartenant au masque de la diapositive en cours, chacune dessinée avec ses vrais éléments graphiques et ses cadres d'espaces réservés plutôt que listée par son nom. Le bouton **Disposition** ouvre la même galerie pour changer la disposition de la diapositive en cours, la disposition actuellement utilisée étant entourée. Changer de disposition déplace votre contenu vers les espaces réservés correspondants ; tout ce pour quoi la nouvelle disposition n'a pas d'emplacement reste sur la diapositive comme contenu libre au lieu d'être supprimé.
- **Réorganiser** - faites glisser les vignettes dans le panneau des diapositives, ou faites glisser les cartes dans la [trieuse de diapositives](/fr/user/viewing#la-trieuse-de-diapositives).
- **Dupliquer, masquer ou supprimer** - ouvrez la trieuse de diapositives, faites un clic droit sur une ou plusieurs diapositives et choisissez **Dupliquer**, **Masquer les diapositives** / **Afficher les diapositives**, ou **Supprimer**. Le copier-coller de diapositives entières y fonctionne aussi (**Ctrl+C** / **Ctrl+V**).
- **Sections** - utilisez **Accueil > Section** ou faites un clic droit sur une vignette et choisissez **Ajouter une section avant** pour regrouper les diapositives sous des titres nommés. Faites un clic droit sur un en-tête de section pour le **Renommer**, le **Supprimer** ou le **Déplacer**.
- **Réinitialiser** - **Accueil > Réinitialiser** restaure les positions des espaces réservés de la diapositive en cours à partir de sa disposition.

## Ajouter une police que la visionneuse ne possède pas

Les présentations utilisent parfois des polices qui ne sont pas installées sur votre
appareil. La visionneuse leur substitue la police disponible la plus proche, ce qui
peut changer l'aspect de la présentation. Si vous avez le fichier de la police, vous
pouvez le confier à la visionneuse pour la session en cours :

1. Ouvrez **Fichier > Options > Général** et repérez la section **Polices**.
2. Activez **M’autoriser à ajouter des fichiers de police à cette session**. Le paramètre
   est désactivé par défaut ; le sélecteur de fichiers reste donc inactif tant que vous
   ne l'activez pas.
3. Cliquez sur **Ajouter un fichier de police** et choisissez un fichier `.ttf`, `.otf`,
   `.woff` ou `.woff2`.

La famille apparaît sous **Polices ajoutées cette session** et dans la liste des polices
de l'onglet Accueil, et les diapositives qui l'utilisent sont réaffichées avec la vraie
police. Les fichiers dont le nom indique une graisse ou un style
(`Inter-SemiBoldItalic.woff2`) s'enregistrent sous le seul nom de famille ; les quatre
fichiers d'une même famille se combinent donc en une seule entrée sur laquelle les
boutons **Gras** et _Italique_ agissent.

Les polices ajoutées ne sont conservées qu'en mémoire. Elles ne sont jamais envoyées
ailleurs, jamais écrites dans la présentation, et sont supprimées au rechargement de la
page : un fichier de police sous licence reste là où vous l'avez placé. Pour livrer une
police _avec_ la présentation, utilisez plutôt **Fichier > Incorporer les polices**.

## Modifier les notes du présentateur

Ouvrez le panneau de notes avec le bouton **Notes** de la barre d'état et saisissez votre texte. La barre d'outils des notes propose **Gras**, **Italique**, **Souligné**, **Barré**, des listes à puces et numérotées, des commandes de retrait et l'insertion de liens, ainsi qu'un commutateur entre l'éditeur de texte riche et brut. Les notes sont enregistrées avec la présentation et apparaissent dans la [vue du présentateur](/fr/user/presenting#vue-du-presentateur).

## Rechercher et remplacer

1. Appuyez sur **Ctrl+F**, ou utilisez **Accueil > Modification > Rechercher** / **Remplacer**.
2. Saisissez le texte à rechercher. Le compteur **x sur y** indique les occurrences dans toutes les diapositives ; parcourez-les avec les boutons précédent / suivant.
3. Activez **Respecter la casse** si nécessaire.
4. Pour modifier du texte, saisissez un remplacement et cliquez sur **Remplacer** ou **Tout remplacer**.

## Commentaires

- Pour commenter un élément, faites un clic droit dessus et choisissez **Ajouter un commentaire** ; ou ouvrez l'onglet **Commentaires** de l'inspecteur et utilisez la zone **Ajouter un commentaire** pour la diapositive en cours.
- Chaque fil permet de **Répondre**, **Modifier**, **Résoudre** (et **Rouvrir**) et **Supprimer**.
- Les diapositives qui comportent des commentaires affichent un badge de commentaire sur leur vignette. Le bouton **Commentaires** de l'onglet **Révision** affiche ou masque le panneau.

## Annuler et rétablir

- **Annuler** - **Ctrl+Z**, ou le bouton flèche de la barre de titre.
- **Rétablir** - **Ctrl+Maj+Z** ou **Ctrl+Y**.

Chaque modification est enregistrée comme une étape que vous pouvez parcourir vers l'arrière et vers l'avant. Les actions continues, comme le glissement, comptent pour une seule étape.

## Couper, copier, coller, dupliquer

| Action    | Raccourci                       |
| --------- | ------------------------------- |
| Copier    | **Ctrl+C**                      |
| Couper    | **Ctrl+X**                      |
| Coller    | **Ctrl+V**                      |
| Dupliquer | **Ctrl+D**                      |
| Supprimer | **Suppr** ou **Retour arrière** |

Les mêmes commandes apparaissent lorsque vous faites un clic droit sur un élément. Voir la liste complète sur la page [Raccourcis clavier](/fr/user/shortcuts).

Pour coller une image depuis une autre application, copiez l'image elle-même, cliquez sur le canevas de la diapositive,
puis appuyez sur **Ctrl+V** (**Cmd+V** sur macOS). Lorsque le navigateur fournit un fichier image,
il est inséré au centre de la diapositive, ajusté à ses limites, et sélectionné.
L'insertion est une seule étape d'annulation. Coller pendant l'édition d'un texte ou d'une cellule de tableau reste
dans cet éditeur ; cela n'insère pas d'image de diapositive distincte. Les URL d'image et le HTML
seuls ne sont pas importés par ce raccourci.

Le presse-papiers interne existant des éléments de l'éditeur reste prioritaire. Si vous avez
copié des éléments de diapositive dans cette visionneuse, **Ctrl+V** continue de coller ces éléments.
Utilisez alors **Insertion > Image** pour choisir un fichier image externe.

## Sauvegarde automatique et récupération

- Le commutateur **Sauvegarde automatique** de la barre de titre active ou désactive les instantanés périodiques. Lorsqu'il est activé, l'application enregistre toutes les deux minutes environ une copie de récupération de votre travail dans le stockage local du navigateur tant que des modifications ne sont pas enregistrées.
- La barre d'état indique l'état actuel : **Modifications non enregistrées**, **Enregistrement...**, **Enregistré [heure]** ou **Tout est enregistré**.
- Si vous fermez l'onglet avec des modifications non enregistrées puis revenez, une bannière intitulée **Modifications non enregistrées récupérées** propose **Restaurer** et **Ignorer**.
- **Historique des versions** (dans le menu de débordement **« ... »**) liste l'instantané de récupération stocké avec son heure et sa taille, et permet de le restaurer ou de le supprimer.

::: warning La sauvegarde automatique n'est pas un enregistrement de fichier
La sauvegarde automatique conserve une copie de récupération uniquement dans votre navigateur. Pour produire un vrai fichier `.pptx`, utilisez le bouton **Enregistrer** de la barre de titre ou **Fichier > Enregistrer** ; le fichier est téléchargé sur votre appareil. Voir [Exportation](/fr/user/exporting).
:::

## Suite

- Présenter votre diaporama : [Présentation](/fr/user/presenting)
- L'enregistrer et le partager : [Exportation](/fr/user/exporting)
- Gagner du temps : [Raccourcis clavier](/fr/user/shortcuts)
