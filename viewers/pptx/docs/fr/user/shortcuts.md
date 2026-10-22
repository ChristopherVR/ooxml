---
title: Raccourcis clavier
description: La référence complète des raccourcis clavier pour l'édition, la navigation, la trieuse de diapositives et les diaporamas.
---

# Raccourcis clavier

Sur macOS, utilisez **Cmd** partout où **Ctrl** est indiqué.

::: info Les raccourcis dépendent du contexte
Les raccourcis d'édition ne fonctionnent qu'en mode édition et sont supprimés pendant que vous saisissez dans une zone de texte, modifiez une cellule de tableau ou utilisez un outil de l'onglet Dessin ; les touches vont alors au texte. Les raccourcis du diaporama ne fonctionnent que pendant la présentation, sauf **F5** / **Maj+F5**, qui en lancent un. **Échap** fonctionne toujours.
:::

::: tip Référence intégrée à l'application
Une liste abrégée des raccourcis est disponible dans l'application : **onglet Aide > Raccourcis clavier**, ou **Raccourcis clavier** dans le menu de débordement **« ... »**.
:::

## Édition

Actifs en mode édition lorsque vous ne saisissez pas dans un élément.

| Action                                                | Raccourci                       |
| ----------------------------------------------------- | ------------------------------- |
| Annuler                                               | **Ctrl+Z**                      |
| Rétablir                                              | **Ctrl+Maj+Z** ou **Ctrl+Y**    |
| Copier l'élément sélectionné                          | **Ctrl+C**                      |
| Couper l'élément sélectionné                          | **Ctrl+X**                      |
| Coller l'élément                                      | **Ctrl+V**                      |
| Dupliquer l'élément sélectionné                       | **Ctrl+D**                      |
| Supprimer l'élément sélectionné                       | **Suppr** ou **Retour arrière** |
| Sélectionner tous les éléments de la diapositive      | **Ctrl+A**                      |
| Déplacer légèrement l'élément sélectionné             | **Touches fléchées**            |
| Déplacer l'élément sélectionné (pas plus grand)       | **Maj+touches fléchées**        |
| Grouper les éléments sélectionnés                     | **Ctrl+G**                      |
| Dissocier les éléments sélectionnés                   | **Ctrl+Maj+G**                  |
| Ouvrir ou fermer Rechercher et remplacer              | **Ctrl+F**                      |
| Ouvrir ou fermer ce panneau de raccourcis             | **?** ou **Ctrl+/**             |
| Désélectionner / fermer les menus / arrêter l'édition | **Échap**                       |

**Ctrl+F** est le seul raccourci qui fonctionne encore pendant que vous saisissez dans un
élément, comme dans PowerPoint : c'est le raccourci qu'on utilise quand le
curseur est déjà dans une zone de texte.

## Édition de texte

Pendant la saisie dans un élément.

| Action                           | Raccourci       |
| -------------------------------- | --------------- |
| Valider la modification du texte | **Ctrl+Entrée** |
| Annuler la modification du texte | **Échap**       |

## Navigation et affichage

| Action                                       | Raccourci                     |
| -------------------------------------------- | ----------------------------- |
| Diapositive précédente (rien de sélectionné) | **Flèche gauche**             |
| Diapositive suivante (rien de sélectionné)   | **Flèche droite**             |
| Zoomer / dézoomer sur le canevas             | **Ctrl+molette de la souris** |

Zoom avant, zoom arrière et ajuster à la fenêtre sont aussi des boutons de la barre d'état.

::: tip
Les touches fléchées ont une double fonction en mode édition : avec un élément sélectionné, elles le déplacent légèrement ; sans sélection, elles changent de diapositive.
:::

## Trieuse de diapositives

Actifs tant que la superposition de la trieuse de diapositives est ouverte.

| Action                                       | Raccourci                       |
| -------------------------------------------- | ------------------------------- |
| Copier les diapositives sélectionnées        | **Ctrl+C**                      |
| Coller des diapositives                      | **Ctrl+V**                      |
| Dupliquer les diapositives sélectionnées     | **Ctrl+D**                      |
| Sélectionner toutes les diapositives         | **Ctrl+A**                      |
| Supprimer les diapositives sélectionnées     | **Suppr** ou **Retour arrière** |
| Zoomer / dézoomer sur les vignettes          | **Ctrl +** / **Ctrl -**         |
| Réduire la sélection, puis fermer la trieuse | **Échap**                       |

::: info Toutes les trieuses n'ont pas toutes les commandes
Les cinq visionneuses répondent à **Échap**, **Suppr** et **Ctrl+D** dans la trieuse. Le presse-papiers de diapositives
(**Ctrl+C** / **Ctrl+V**), **Ctrl+A** et les touches de zoom des vignettes nécessitent une trieuse à sélection multiple
et zoomable, que seule la visionneuse React fournit actuellement ; ailleurs, ces combinaisons sont transmises au
navigateur au lieu de ne rien faire en silence.
:::

## Diaporama

Le lancement d'un diaporama fonctionne depuis n'importe où dans la visionneuse, y compris pendant la saisie, et dans une
visionneuse en lecture seule également ; les autres raccourcis ne sont actifs que pendant la présentation.

| Action                                       | Raccourci  |
| -------------------------------------------- | ---------- |
| Lancer le diaporama depuis le début          | **F5**     |
| Lancer le diaporama depuis cette diapositive | **Maj+F5** |

::: warning F5 ne recharge plus la page lorsque la visionneuse a le focus
Ce sont les touches de PowerPoint, et la touche de rechargement du navigateur. La visionneuse ne s'approprie un simple **F5**
que lorsque le focus clavier est à l'intérieur et qu'aucun diaporama n'est en cours ; **Ctrl+F5** et le bouton
de rechargement du navigateur ne sont pas touchés.
:::

Ils correspondent aux touches de diaporama propres à PowerPoint. Notez la séparation délibérée : les outils
d'annotation sont des combinaisons **Ctrl**, car PowerPoint réserve les lettres seules `N`, `P`, `B`, `W`
et `E` à la navigation et aux commandes d'écran.

| Action                                            | Raccourci                                                                             |
| ------------------------------------------------- | ------------------------------------------------------------------------------------- |
| Diapositive suivante / étape d'animation suivante | **Flèche droite**, **Flèche bas**, **Page suivante**, **Espace**, **Entrée** ou **N** |
| Diapositive précédente                            | **Flèche gauche**, **Flèche haut**, **Page précédente**, **Retour arrière** ou **P**  |
| Première diapositive                              | **Début**                                                                             |
| Dernière diapositive                              | **Fin**                                                                               |
| Aller à une diapositive                           | Tapez son numéro, puis **Entrée**                                                     |
| Écran noir activé / désactivé                     | **B** ou **.**                                                                        |
| Écran blanc activé / désactivé                    | **W** ou **,**                                                                        |
| Pointeur flèche (normal)                          | **Ctrl+A**                                                                            |
| Pointeur laser                                    | **Ctrl+L**                                                                            |
| Stylo                                             | **Ctrl+P**                                                                            |
| Gomme                                             | **Ctrl+E**                                                                            |
| Effacer toutes les annotations d'une diapositive  | **E**                                                                                 |
| Afficher / masquer les annotations à l'encre      | **Ctrl+M**                                                                            |
| Afficher / masquer la barre d'outils du diaporama | **Ctrl+H**                                                                            |
| Afficher toutes les diapositives                  | **Ctrl+S**                                                                            |
| Sous-titres en direct activés / désactivés        | **J**                                                                                 |
| Menu contextuel du diaporama                      | **Maj+F10** ou la touche **Menu**                                                     |
| Terminer le diaporama                             | **Échap** ou **-**                                                                    |

::: info Prise en charge des deux combinaisons d'interface
**Ctrl+H** (masquer la barre d'outils du diaporama) et **Ctrl+S** (Afficher toutes les diapositives) sont reconnues par la
table de raccourcis de diaporama partagée et gérées par chaque visionneuse.
:::

Le surligneur n'a pas de raccourci clavier ; utilisez la barre d'outils du diaporama. La vue du présentateur
est elle aussi un commutateur de la barre d'outils, pas une touche. Voir [Présentation](/fr/user/presenting) pour le
fonctionnement des outils d'annotation et de la vue du présentateur.
