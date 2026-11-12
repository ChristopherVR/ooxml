---
title: Environnements d'exécution
description: Où s'exécute pptx-viewer (navigateur, Node.js, Web Worker) et le comportement propre à la plateforme qui découle de l'exécution dans le bac à sable d'un navigateur.
---

# Environnements d'exécution

`pptx-viewer` se divise en un moteur principal sans DOM et en liaisons d'interface réservées au navigateur. Cette page décrit où s'exécute chaque partie, ainsi que le comportement qui découle directement du bac à sable du navigateur plutôt que d'une fonctionnalité absente.

## Où il s'exécute {#where-it-runs}

| Environnement           | Fonctionne  | Restrictions                                                                                                                                                                                                                      |
| ----------------------- | ----------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Navigateur              | Oui         | L'ensemble complet des fonctions : analyse, rendu, édition, export, collaboration.                                                                                                                                                |
| Node.js (et serverless) | Moteur seul | `pptx-viewer-core` (chargement, édition, enregistrement, conversion Markdown/SVG, chiffrement) est sans DOM. Les liaisons d'interface, l'export raster (`html2canvas`) et la conversion EMF/WMF sont des fonctions du navigateur. |
| Web Worker              | Moteur seul | Même périmètre que Node.js : le moteur ne dépend d'aucun DOM.                                                                                                                                                                     |

## Remarques sur la plateforme {#platform-notes}

Ce ne sont pas des lacunes à combler : elles découlent directement de l'exécution dans le bac à sable de sécurité d'un navigateur, qui ne peut pas faire ce que peut faire l'application de bureau.

### Polices {#fonts}

Un navigateur n'a pas de catalogue de polices propre sur lequel s'appuyer, contrairement à l'application de bureau. `pptx-viewer` résout donc chaque police de la présentation en trois étapes : l'utiliser telle quelle lorsqu'elle est installée sur la machine du lecteur ; sinon, télécharger son clone vérifié à métriques compatibles depuis Google Fonts avant la mise en page (Calibri devient Carlito, Cambria devient Caladea, Arial devient Arimo, Times New Roman devient Tinos, Courier New devient Cousine, Georgia devient Gelasio, de sorte que les sauts de ligne et les dimensions du texte restent conformes à PowerPoint) ; et seulement si aucune des deux n'existe, revenir par la classification PANOSE vers une police générique de même classe, ce qui peut modifier les métriques.

Pendant le chargement d'une présentation, la visionneuse attend brièvement (150 ms) les polices Google Fonts utilisées par ses diapositives, afin que les polices déjà en cache n'apparaissent jamais dans une version de remplacement au premier affichage. Un premier téléchargement sur un réseau lent ne retient pas la présentation : après 150 ms, ou immédiatement si la requête échoue (hors ligne, bloquée), la présentation s'affiche avec des substituts à métriques compatibles, la barre d'état indique « Téléchargement des polices (1 sur 2) » pendant le téléchargement, et une police arrivée en retard est remise en page à son arrivée. Les polices installées et intégrées n'imposent aucune attente.

### Polices et confidentialité {#fonts-and-privacy}

Récupérer un clone depuis Google Fonts signifie que le navigateur du lecteur contacte `fonts.googleapis.com` et `fonts.gstatic.com`, ce qui révèle l'adresse IP du lecteur à Google. Par défaut, c'est le lecteur qui décide avec l'option « Télécharger les polices non installées ni intégrées (Google Fonts) » du menu Fichier > Options > Trust Center (Centre de gestion de la confidentialité), qui est activée. La désactiver s'applique immédiatement à la présentation ouverte (la visionneuse supprime sa feuille de style de polices web) et à chaque chargement ultérieur, et le choix est enregistré avec les autres options. Un hôte peut décider à la place avec l'option `googleFonts`, disponible sur toutes les liaisons :

- non défini (par défaut) : l'option du Trust Center décide ; lorsqu'elle est activée, les clones sont chargés depuis Google Fonts comme décrit ci-dessus ;
- `googleFonts: false` : aucune requête ne part vers Google ; le texte utilise les polices installées, les polices intégrées à la présentation, ou le repli PANOSE, de sorte que les sauts de ligne peuvent différer de PowerPoint. L'option du Trust Center apparaît désactivée et en lecture seule ;
- `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }` : la même requête CSS2 part vers un miroir auto-hébergé à la place, et l'option du Trust Center apparaît activée et en lecture seule. Le miroir doit répondre au format de requête CSS2 de Google Fonts (`?family=...&display=swap`) avec un CSS `@font-face` compatible, et ses fichiers de polices doivent aussi être auto-hébergés pour qu'aucune requête n'atteigne Google. Un `baseUrl` qui n'est pas une URL `http(s)` absolue désactive le repli (l'option apparaît alors désactivée et en lecture seule) et journalise un avertissement.

L'option gouverne aussi la récupération des contours de glyphes de WordArt et les polices web intégrées aux exports d'images de diapositives, puisque les deux réutilisent l'URL de feuille de style résolue par la visionneuse. Les deux lisent le CSS et les fichiers de polices avec `fetch()`, donc un miroir doit envoyer des en-têtes CORS (`Access-Control-Allow-Origin`) pour fonctionner, comme le font les serveurs de Google.

La manière dont la visionneuse choisit une police pour chaque segment et attend les téléchargements est décrite dans le [chargement des polices](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md). Chaque requête que les packages peuvent émettre est listée dans [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

### Contenu externe et confidentialité {#external-content-and-privacy}

Une présentation peut lier des images (y compris les remplissages et les puces en image), des vidéos, des audios et des images d'affiche de cadre vidéo vers d'autres serveurs (relations `TargetMode="External"`) au lieu de les intégrer. À l'ouverture d'une telle présentation, le navigateur du lecteur va chercher ces ressources, ce qui révèle à l'exploitant du serveur l'adresse IP du lecteur et le fait que la présentation a été ouverte. Par défaut, c'est l'option « Autoriser le contenu externe » du Trust Center qui décide, et elle est activée. Un hôte peut décider à la place avec l'option `allowExternalContent`, disponible sur toutes les liaisons :

- non défini (par défaut) : l'option du Trust Center décide ;
- `allowExternalContent: false` : aucun contenu lié n'est récupéré, de sorte que les diapositives n'affichent que le contenu intégré, et l'option du Trust Center apparaît désactivée et en lecture seule ;
- `allowExternalContent: true` : le contenu lié est récupéré, et l'option apparaît activée et en lecture seule.

La valeur s'applique à chaque présentation ouverte par la visionneuse, y compris celles ouvertes pour la Comparaison et les diapositives réutilisées depuis un autre fichier ; un changement prend effet au chargement suivant. Les liens hypertexte ne sont pas concernés : ils ne s'ouvrent que lorsque le lecteur les suit. Les polices sont distinctes : les polices de la présentation qui ne sont ni installées ni intégrées se chargent depuis Google Fonts, sauf si l'option du Trust Center « Télécharger les polices non installées ni intégrées » ou l'option `googleFonts` les désactive (voir [Polices et confidentialité](#fonts-and-privacy)).

Dans une session de collaboration, les diapositives arrivent déjà analysées depuis les pairs, et le paramètre ne leur est pas appliqué une seconde fois : un pair dont la visionneuse autorise le contenu externe peut partager une diapositive qui référence une URL liée. Utilisez la même valeur `allowExternalContent` pour tous les participants si cela compte.

### Limites de chargement {#load-limits}

Un hôte qui ouvre des présentations qu'il ne juge pas fiables (envois publics, pièces jointes) peut borner le travail qu'une présentation peut provoquer avec l'option `limits`, disponible sur toutes les liaisons et transmise telle quelle à `PptxHandler.load` :

| Champ                  | Valeur par défaut | Bornes                                                                                                                               |
| ---------------------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `maxInputBytes`        | aucune limite     | Le fichier tel que donné à la visionneuse, vérifié avant son ouverture.                                                              |
| `maxUncompressedBytes` | 500 MiB           | L'archive une fois décompressée, vérifiée d'après les tailles déclarées puis de nouveau pendant la lecture de chaque partie.         |
| `maxEntries`           | 65,536            | Les parties (entrées ZIP) de l'archive ; 65,536 est aussi la valeur la plus élevée acceptée.                                         |
| `maxSlides`            | aucune limite     | Les diapositives listées dans `ppt/presentation.xml`, vérifiées avant l'analyse de la moindre diapositive.                           |
| `maxShapes`            | aucune limite     | Les éléments créés sur les diapositives (membres de groupes compris), comptés au fur et à mesure de l'analyse de chaque diapositive. |

Si l'option est absente ou si un champ est omis, la valeur par défaut est conservée : une visionneuse sans `limits` se comporte comme auparavant. Une présentation qui dépasse une limite cesse de se charger à cette vérification, et la visionneuse affiche son écran d'erreur de chargement avec un message localisé qui nomme la limite et la valeur de la présentation, par exemple « Cette présentation compte 120 diapositives, au-delà des 100 acceptées par cette visionneuse. » Une valeur qui n'est pas un nombre positif fini est rejetée comme erreur au lieu d'être ignorée. Les limites s'appliquent à la présentation ouverte par la visionneuse (les `content` / `source` et Fichier > Ouvrir), et non à un fichier que l'utilisateur choisit pour la Comparaison ou la Réutilisation de diapositives ; un changement prend effet au chargement suivant. Sans la visionneuse, passez le même objet sous la forme `handler.load(buffer, { limits })` (voir [Options de chargement](/core/loading#load-options)).

### Exécuter des actions de programme (`ppaction://program`) {#run-program-actions-ppaction-program}

L'action « Exécuter un programme » de PowerPoint est analysée et réécrite sans perte (le chemin ou la commande de programme saisi survit au chargement et à l'enregistrement sans modification), et la boîte de dialogue Paramètres de l'action permet de la consulter et de la modifier comme toute autre action pendant l'édition. Ce qu'un navigateur ne peut pas faire, c'est ce que PowerPoint fait lui-même lorsque l'action se déclenche pendant un diaporama : lancer un exécutable local. Aucune API de plateforme web n'existe pour cela, par conception (une page capable de lancer des programmes locaux arbitraires serait une faille de sécurité critique).

À la place, un clic sur une forme Exécuter un programme pendant un diaporama affiche une notification non bloquante qui nomme la commande résolue exacte (le chemin et les arguments que l'auteur a saisis ; OOXML les stocke comme une seule chaîne opaque, il n'y a rien à découper), avec un bouton Copier en un clic. Le présentateur voit ainsi ce que PowerPoint aurait exécuté et peut le lancer lui-même en dehors du navigateur, s'il le souhaite. Le clic compte toujours comme traité (le diaporama n'avance pas en plus), comme pour tous les autres types d'action. Cela est implémenté une seule fois dans une fonction de décision partagée (l'intention `runProgram` de `src/ui/src/pptx/render/presentation-action.ts`, ainsi que `src/ui/src/pptx/render/run-program-notice.ts`), utilisée à l'identique par les cinq liaisons, et `e2e/pptx/run-program-notice.spec.ts` teste la notification, le bouton Copier et la consommation du clic sur les cinq démonstrations.

Deux vrais bogues sont apparus et ont été corrigés pendant ce travail : React affichait initialement la notification comme frère de l'élément sur lequel le diaporama appelle `requestFullscreen()`, et la couche supérieure d'un élément en plein écran réel du système s'affiche au-dessus de tout ce qui est hors de lui, quel que soit le `z-index` CSS, de sorte que le bouton Copier était visible mais non cliquable ; elle s'affiche maintenant par l'emplacement de superposition de la scène du diaporama, aux côtés des autres éléments d'interface réservés au diaporama. Par ailleurs, une cible qui n'a pas la forme d'une URI (un chemin saisi comme `notepad.exe C:\temp\notes.txt`) n'était pas reconnue comme cible de relation externe à l'enregistrement, de sorte que le paquet écrit échouait la validation `TargetMode="External"` pour `ppaction://program`, `hlinkfile` et `hlinkpres` ; le chemin d'enregistrement force maintenant `External` pour les trois.

### Lecture des médias {#media-playback}

La lecture audio et vidéo dépend de la prise en charge des codecs par le navigateur lui-même : les WMV et autres codecs anciens peuvent ne pas se lire, et les médias protégés par DRM ne se liront pas. C'est un fait de la plateforme, non une fonctionnalité absente : `pptx-viewer` confie la lecture à l'élément multimédia natif du navigateur, et un navigateur ne décode que ce pour quoi il embarque un codec.

## Lectures associées {#related-reading}

- [Limitations](/fr/guide/limitations) - lacunes ouvertes et non résolues.
- [Fidélité des effets visuels](/fr/guide/visual-effects) - approximations CSS/SVG des effets et les preuves qui les appuient.
