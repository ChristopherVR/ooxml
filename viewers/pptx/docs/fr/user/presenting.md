---
title: Présentation
description: Lancez un diaporama plein écran, pilotez-le au clavier, utilisez la vue du présentateur, annotez avec le stylo et le surligneur et répétez les minutages.
---

# Présentation

Le mode présentation transforme la visionneuse en un diaporama plein écran qui joue vos transitions et animations, avec une vue du présentateur, des minuteurs et des outils d'annotation en direct.

## Lancer un diaporama

Depuis l'onglet **Diaporama** :

1. Cliquez sur **Depuis le début** pour commencer à la diapositive 1, ou sur **À partir de la diapositive actuelle** pour commencer là où vous êtes.
2. La diapositive s'agrandit pour remplir l'écran et l'interface d'édition disparaît.

Vous pouvez aussi cliquer sur le bouton **Présenter** de la rangée supérieure de la barre d'outils, ou sur le bouton **Diaporama** de la barre d'état. La flèche à côté de **Présenter** ouvre un menu avec **Vue du présentateur**, **Enregistrer les minutages**, **Paramètres du diaporama**, **Présenter en ligne** et un commutateur **Sous-titres**.

![Mode diaporama en plein écran](/user-guide/presenting-slideshow.jpg)

**Paramètres du diaporama** contrôle des options de lecture comme **Utiliser les minutages** (avance automatique selon les minutages enregistrés des diapositives) et la lecture en boucle du diaporama jusqu'à ce que vous appuyiez sur Échap.

## Contrôler le diaporama

| Pour...                                            | Faites ceci                                                                                   |
| -------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| Diapositive suivante / étape d'animation suivante  | Cliquez, ou appuyez sur **Flèche droite**, **Page suivante**, **Espace**, **Entrée** ou **N** |
| Diapositive précédente                             | Appuyez sur **Flèche gauche**, **Page précédente**, **Retour arrière** ou **P**               |
| Aller à une diapositive précise                    | Tapez son numéro, puis appuyez sur **Entrée**                                                 |
| Première / dernière diapositive                    | Appuyez sur **Début** / **Fin**                                                               |
| Obscurcir l'écran (noir)                           | Appuyez sur **B** (appuyez de nouveau pour reprendre)                                         |
| Éclaircir l'écran (blanc)                          | Appuyez sur **W** (appuyez de nouveau pour reprendre)                                         |
| Basculer la vue du présentateur                    | Utilisez le bouton **Vue du présentateur** de la barre d'outils (pas de raccourci clavier)    |
| Afficher ou masquer la barre d'outils du diaporama | Appuyez sur **Ctrl+H**                                                                        |
| Afficher toutes les diapositives                   | Appuyez sur **Ctrl+S**                                                                        |
| Terminer le diaporama                              | Appuyez sur **Échap**                                                                         |

### La barre d'outils du diaporama

Une barre d'outils flottante se trouve en bas de l'écran et comporte, de gauche à droite : **Diapositive précédente**, un compteur de diapositives (par exemple `3 / 12`), **Diapositive suivante**, un minuteur écoulé, **Pointeur laser**, **Stylo**, **Surligneur**, **Gomme**, **Effacer les annotations**, **Vue du présentateur** et **Terminer la présentation**.

![Barre d'outils du diaporama](/user-guide/presenting-toolbar.jpg)

::: tip Masquage automatique
La barre d'outils s'estompe après environ 3 secondes sans mouvement de la souris. Bougez la souris (surtout vers le bas de l'écran) pour la faire réapparaître, ou appuyez sur **Ctrl+M** pour l'épingler ou la libérer.
:::

## Transitions et animations

Lorsque vous avancez d'une diapositive à l'autre, la **transition** assignée à la diapositive entrante se joue : fondus, poussées, essuyages, révélations, morphes, etc. Le retour en arrière est instantané.

![Diaporama avec transitions entre les diapositives](/user-guide/presenting-transitions.gif)

Les **animations** se jouent selon leurs déclencheurs :

- **Au clic** - chaque clic ou touche d'avance joue l'étape d'animation suivante ; la diapositive ne change qu'une fois toutes les étapes jouées.
- **Avec la précédente** / **Après la précédente** - ces étapes s'enchaînent automatiquement après celle qui les précède.

Les diapositives dotées de minutages enregistrés avancent automatiquement lorsque **Utiliser les minutages** est activé dans Paramètres du diaporama. Les éléments avec des déclencheurs au survol ou au clic (y compris les boutons d'action et les liens hypertexte) réagissent pendant le diaporama.

::: warning Notes sur la fidélité
Les transitions et animations sont des approximations du comportement de PowerPoint. Morphe interpole la position, la taille, l'opacité, la rotation et la couleur des éléments correspondants, et assure aussi un morphe entre des géométries de forme différentes : lorsqu'un élément apparié change de type de forme (ou de contour piloté par des ajustements) entre les deux diapositives, son contour est converti en polygone et interpolé sommet par sommet vers le contour de la forme cible, au lieu d'un fondu enchaîné. Voir [Limitations](/fr/guide/limitations) pour les autres approximations de transitions et d'animations.

**Les éléments non appariés font un fondu enchaîné, comme dans PowerPoint.** Les éléments sont appariés entre les deux diapositives par un nommage explicite avec `!!`, par l'identité `a16:creationId` / identifiant de forme natif, et par proximité de même type ; un simple nom du volet Sélection sans préfixe `!!` n'est pas non plus un signal d'appariement que respecte le Morphe de PowerPoint, et renommer deux formes avec le même nom ne les fait donc pas passer en morphe. Un élément sans équivalent sur la diapositive suivante disparaît en fondu enchaîné au lieu de passer en morphe. Confirmé sur le vrai PowerPoint 2016 (COM `CreateVideo`, 30 i/s, une présentation de deux diapositives avec un élément non apparié uniquement sur la diapositive 1) : l'élément reste à pleine opacité jusqu'à la dernière image du temps d'affichage de la diapositive 1, puis s'estompe progressivement vers l'arrière-plan en 10 images (environ 0,45 s des 2 s de transition définies), pour devenir totalement transparent, au lieu de disparaître en une seule image.
:::

## Vue du présentateur

La vue du présentateur vous montre des informations que le public ne voit pas :

- La diapositive en cours (avec vos outils d'annotation toujours disponibles).
- Un aperçu **Diapositive suivante**.
- Les **Notes du présentateur**, avec des commandes de taille de police ; vous pouvez y modifier directement les notes.
- L'heure actuelle, un minuteur écoulé avec pause et réinitialisation, et une barre de progression.

Activez la vue du présentateur avec le bouton **Vue du présentateur** de la barre d'outils du diaporama, ou depuis **Diaporama > Vue du présentateur** dans le ruban. Il n'existe pas de raccourci clavier pour cela : `N` est la touche « diapositive suivante » de PowerPoint.

![Vue du présentateur affichant les notes, l'aperçu de la diapositive suivante et le minuteur](/user-guide/presenting-presenter-view.jpg)

La barre d'outils supérieure de la console du présentateur propose aussi : pause/réinitialisation du minuteur, une grille **Diapositives** pour accéder à n'importe quelle diapositive, des commandes de zoom, les outils de pointeur, des boutons d'obscurcissement **B** / **W**, un commutateur de sous-titres et un bouton **Affichage du public** qui ouvre le diaporama dans une seconde fenêtre ; utilisez **Permuter** pour changer ce que montre chaque écran, afin que le public ne voie que les diapositives pendant que vous gardez la console.

## Annoter en direct

Dessinez sur les diapositives pendant la présentation pour souligner des points :

| Outil                          | Ce qu'il fait                                                      | Raccourci                   |
| ------------------------------ | ------------------------------------------------------------------ | --------------------------- |
| **Flèche**                     | Le pointeur normal ; ne laisse aucune trace.                       | **Ctrl+A**                  |
| **Pointeur laser**             | Un point lumineux qui suit votre curseur ; ne laisse aucune trace. | **Ctrl+L**                  |
| **Stylo**                      | De l'encre à main levée de la couleur de votre choix.              | **Ctrl+P**                  |
| **Surligneur**                 | Traits de surbrillance translucides.                               | (barre d'outils uniquement) |
| **Gomme**                      | Supprime les traits d'encre un par un.                             | **Ctrl+E**                  |
| **Effacer les annotations**    | Supprime toute l'encre de la diapositive en cours.                 | **E**                       |
| **Afficher / masquer l'encre** | Masque l'encre sans la supprimer.                                  | **Ctrl+M**                  |

::: tip Pourquoi les outils sont des combinaisons Ctrl
PowerPoint réserve les lettres seules à la navigation et aux commandes d'écran (`N` suivant,
`P` précédent, `B` noir, `W` blanc, `E` tout effacer) ; les outils d'annotation prennent donc la
combinaison **Ctrl**. Cette visionneuse suit la même répartition.
:::

Pour changer de couleur, faites un clic droit sur le bouton du stylo ou du surligneur (ou cliquez sur la petite flèche à côté) et choisissez dans la grille d'échantillons. Choisir une couleur active aussi cet outil.

![Annotation au stylo sur une diapositive pendant la présentation](/user-guide/presenting-pen-annotation.gif)

Les annotations sont conservées par diapositive : passez à une autre diapositive puis revenez, et votre encre est toujours là.

![Annotation dessinée sur une diapositive](/user-guide/presenting-annotation.jpg)

Lorsque vous terminez un diaporama qui comporte de l'encre, une boîte de dialogue **Conserver les annotations ?** demande s'il faut garder les dessins (ils deviennent des éléments d'encre sur les diapositives) ou les abandonner.

## Répéter avec les minutages

1. Cliquez sur **Diaporama > Enregistrer les minutages**. Le diaporama démarre avec un petit panneau de minuterie indiquant **Durée de la diapositive** et **Durée totale** ; un bouton **Pause** arrête le chronomètre pendant une interruption.
2. Répétez votre exposé en avançant comme vous le feriez normalement. Le temps passé sur chaque diapositive est enregistré.
3. Appuyez sur **Échap** (ou avancez au-delà de la dernière diapositive). Un récapitulatif **Minutages de la répétition** liste chaque diapositive avec sa durée et votre total.
4. Cliquez sur **Enregistrer les minutages** pour conserver les durées par diapositive (utilisées pour l'avance automatique lorsque **Utiliser les minutages** est activé), ou sur **Abandonner**.

## Présenter à des spectateurs distants

**Diaporama > Diffuser** (aussi **Présenter en ligne** dans le menu Présenter) démarre une session à sens unique où les spectateurs suivent vos diapositives en direct à partir d'un lien. Voir [Collaboration](/fr/user/collaboration#diffuser-un-diaporama).

## Quitter

Appuyez sur **Échap** ou cliquez sur le **X** de la barre d'outils du diaporama. Si vous avez des annotations, la boîte de dialogue conserver-ou-abandonner apparaît ; si vous répétiez, le récapitulatif des minutages apparaît.

## Suite

- Enregistrer un document ou un enregistrement de votre présentation : [Exportation](/fr/user/exporting)
- Liste complète des touches : [Raccourcis clavier](/fr/user/shortcuts)
