---
title: Architecture
description: Comment pptx-viewer est structuré - les pipelines de chargement et de sauvegarde, le runtime composé par mixins, la résolution des thèmes, le moteur de géométrie et la couche de rendu partagée qui alimente chaque binding.
---

# Architecture

`pptx-viewer` est un système en couches. Les bindings de framework (React, Vue 3, Angular, Svelte et un binding JavaScript natif en DOM pur) gèrent l'interface ; une couche de rendu partagée fournit la logique indépendante du framework ; et le moteur Core gère tout ce qui touche à l'analyse, l'édition et la sauvegarde des fichiers PowerPoint. Chaque couche ne dépend que de celle du dessous.

## Aperçu

```
+---------------------------------------------------------------+
|                      Framework bindings                       |
|   React    |   Vue 3   |  Angular  |  Svelte 5  |  Vanilla JS |
| pptx-react | pptx-vue  | pptx-ang. | pptx-svelte| pptx-vanilla|
+------------------------------+--------------------------------+
                               |
                               v
+---------------------------------------------------------------+
|             Shared rendering layer (ooxml-ui/pptx)            |
|     geometry, styles, gradients, charts, connectors, text     |
+------------------------------+--------------------------------+
                               |
                               v
+---------------------------------------------------------------+
|                Core engine (pptx-viewer-core)                 |
|  PptxHandler (public API)                                     |
|    -> Runtime (parsing, serialization, theme resolution)      |
|         -> Types, Geometry, Colour, Builders, Converter       |
+---------------------------------------------------------------+
```

## Bindings de framework

Les packages de bindings sont de fines couches de présentation. Ils consomment les données de rendu précalculées de `ooxml-ui/pptx` et les traduisent en modèles propres à chaque framework (JSX, SFC Vue, composants Angular, runes Svelte 5 ou appels DOM natifs pour le binding vanilla). Les diapositives sont rendues en HTML/SVG mis à l'échelle par des transformations CSS, ce qui donne un texte net à tout zoom, une accessibilité native et une interactivité DOM complète.

Chaque binding expose un point d'entrée de visionneuse/éditeur de haut niveau qui orchestre l'état, l'édition, le chargement, l'export et le mode présentation selon l'idiome natif de son framework : hooks dans React, composables dans Vue, services dans Angular, runes dans Svelte, et une simple fonction de fabrique avec une API d'instance impérative en JavaScript natif. Comme tous consomment la même couche partagée, le rendu est identique dans les cinq.

## Couche de rendu partagée (`ooxml-ui/pptx`)

Le rendu partagé vit dans `src/ui/src/pptx/` à la racine du dépôt et est publié via `ooxml-ui/pptx`. Les opérations de document sans DOM vivent dans `src/core/pptx/editor/`, et les schémas d'outils dans `src/core/pptx/automation/schemas/`. Le moteur de rendu comprend :

- **Routage des connecteurs** : un routeur A\* sur un graphe d'obstacles (`connector-router-astar.ts`, `connector-router-graph.ts`, désormais dans l'aire partagée `ooxml-core/geometry`) ainsi que la construction et le re-routage des tracés.
- **Mathématiques des graphiques** : plages d'axes, positionnement des catégories, constructeurs de tracés cartésiens/polaires, statistiques de boîtes à moustaches, composition combinée/boursière (la famille de modules `chart-*`).
- **Moteurs d'animation et de morphe** : construction et lecture de la chronologie (`animation-timeline-*.ts`), ainsi que l'appariement des transitions morphe et l'interpolation de géométrie (`morph-matching.ts`, `morph-geometry-interp.ts`, `morph-text.ts`).
- **Résolution du texte et des styles** : numérotation des puces, styles de remplissage/dégradé, filtres d'effets d'image, styles de coupure de ligne kinsoku.
- **Mathématiques** : convertisseurs OMML vers MathML et LaTeX vers OMML (`omml-to-mathml.ts`, `latex-to-omml.ts`, désormais dans l'aire partagée `ooxml-core/math`).
- **Comportement de l'éditeur** : presse-papiers, repères d'alignement, reproduction de la mise en forme, rechercher et remplacer, synchronisation de la collaboration et présence.

Les répertoires voisins couvrent `export/`, `i18n/`, `loader/`, `theme/` et le moteur de rendu optionnel `smartart-3d/`.

### Contrôleurs sans interface

La logique d'éditeur avec état est écrite une seule fois comme contrôleur sans interface dans `render/controller/` : un objet simple avec `getSnapshot()`, `subscribe(listener)` et des actions typées. L'historique annuler et rétablir (profondeur par défaut de 100 étapes), la sauvegarde automatique de récupération (un seul moteur, qui capture aussi le texte en ligne non validé), l'offre de récupération, le cycle de vie de la salle de collaboration et le volet IA sont des contrôleurs, de même que les gestes sur la zone de travail (`createGestureController`). Le répartiteur de touches (`dispatchEditorKey`), le pont IA (`createPptxAiBridge`), les cibles de clic sur un groupe et les étapes du diaporama sont des fonctions pures à côté. Chaque liaison raccorde son propre état réactif à un adaptateur de magasin générique (`useController` dans React et Vue, `controllerState` dans Svelte, `controllerSignal` dans Angular, `subscribe` en JavaScript natif), de sorte qu'une correction d'un contrôleur atteint les cinq liaisons.

Les sections Effets, Ligne, Taille et position, et Zone de texte de l'inspecteur sont des descripteurs de champs dans `render/inspector/`, rendus par un seul moteur de rendu générique par liaison. La feuille de données du graphique, le volet Format du graphique, les boîtes de dialogue d'insertion de graphique et de SmartArt, le volet texte SmartArt et les galeries de création de graphiques sont des éléments `ooxml-ui` partagés que chaque liaison monte.

::: info Packages publics partagés
Chaque binding déclare `ooxml-ui` comme dépendance d'exécution et importe ses entrées PowerPoint. Angular utilise la même API publique sans embarquer de copie des sources. Le package privé `pptx-viewer-shared` reste une façade de compatibilité et ne contient aucune implémentation.
:::

## Moteur Core (`pptx-viewer-core`)

Le package Core est entièrement indépendant du framework. Il s'exécute dans tout environnement JavaScript : navigateur, Node.js, Web Worker ou fonction serverless. Son point d'entrée public est `PptxHandler`.

> **Où se trouve le code.** Le moteur est l'aire `pptx` du package public partagé `ooxml-core` (`src/core/pptx/`, sous-chemins `/pptx`, `/pptx/converter`, `/pptx/cli`, `/pptx/signature-node`) ; le dossier de la visionneuse conserve les adaptateurs de framework. `pptx-viewer-core` est un fin package qui réexporte ces sous-chemins avec une API publique inchangée ; tout chemin ci-dessous commençant par `src/core/pptx/` est donc relatif à la racine du monorepo. Le moteur utilise encore son propre modèle d'objets `fast-xml-parser` et est compilé avec des options TypeScript assouplies ; unifier son modèle XML avec l'aire partagée `xml` et durcir les options sont les prochaines étapes de la migration. Les primitives de couleur, la géométrie des formes prédéfinies et les constantes EMU sont partagées avec les autres aires Office (`color`, `geometry`, `units`).

### La façade et le runtime composé par mixins

Le moteur est construit en trois couches de surface décroissante :

```
PptxHandler                    static factories (create / createBlank)
  └─ PptxHandlerCore           thin facade: load / save / export / encryption
       └─ IPptxHandlerRuntime  the actual engine, assembled from many mixin modules
```

- **`PptxHandler`** (`src/core/pptx/core/PptxHandler.ts`) ajoute les points d'entrée de construction statiques `create()` / `createBlank()`.
- **`PptxHandlerCore`** délègue tout l'analyse lourde, la sérialisation et la manipulation XML à un `IPptxHandlerRuntime` injecté. Le runtime est remplaçable via les dépendances du constructeur (`runtime` ou `runtimeFactory`), ce qui permet aux tests et aux hôtes alternatifs d'échanger les implémentations.
- **`PptxHandlerRuntime`** n'est pas une seule classe dans un seul fichier. Il est composé de nombreux **modules ciblés** dans `src/core/pptx/core/core/runtime/`, chacun nommé `PptxHandlerRuntime<Concern>.ts` et traitant exactement une préoccupation : `PptxHandlerRuntimeChartParsing.ts`, `PptxHandlerRuntimeThemeLoading.ts`, `PptxHandlerRuntimeSaveElementWriter.ts`, `PptxHandlerRuntimeSmartArtParsing.ts`, etc.

Chaque module déclare une classe qui étend la classe exportée par le module précédent, formant une chaîne d'héritage linéaire qui empile les capacités :

```ts
// PptxHandlerRuntimeLoadPipeline.ts
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeLoadSession';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	// adds the load-pipeline capability on top of everything below it
}
```

Les nouvelles capacités du moteur (y compris les nouveaux types d'éléments) s'ajoutent comme de nouveaux maillons de cette chaîne plutôt qu'en grossissant les fichiers existants. Les collaborateurs transverses (avertissements de compatibilité, fabriques XML, constructeurs de types de contenu) sont définis derrière des interfaces de service `I*` et injectés par une fabrique de dépendances, ce qui garde chaque mixin testable isolément.

### Le pipeline de chargement

`handler.load(arrayBuffer)` parcourt ces étapes :

```
ArrayBuffer
  │  container sniff: an OLE compound file is either a legacy binary
  │  .ppt (converted to an in-memory .pptx package) or an encrypted
  │  OOXML package (decrypted with the supplied password)
  │  ZIP signature check on whatever comes out
  ▼
JSZip.loadAsync                 in-memory archive
  │  zip-bomb guard: 500 MiB uncompressed budget (configurable),
  │  hard cap of 65,536 entries
  ▼
fast-xml-parser                 XML parts -> JS object trees
  │  Strict OOXML detection: strict namespace URIs are transparently
  │  normalized to Transitional for all subsequent parses
  ▼
Theme / master / layout resolution
  │  colour maps, font schemes, format schemes, placeholder styles
  ▼
Per-slide shape-tree (spTree) parsing
  │  each <p:sp>, <p:pic>, <p:graphicFrame>, ... becomes a typed PptxElement
  ▼
PptxLoadDataBuilder             assembles the final model
  ▼
PptxData
```

1. **Contrôles du conteneur d'abord.** Un fichier composé OLE n'est pas forcément une erreur : c'est ainsi qu'arrivent à la fois un `.ppt` binaire historique et un package OOXML chiffré. Une présentation PowerPoint 97-2003 (reconnue à son flux `PowerPoint Document`) est convertie par `core/ppt/` en un package `.pptx` équivalent en mémoire et chargée à partir de là, de sorte que le reste du pipeline ne sait jamais que la présentation était binaire. Sinon, le conteneur est traité comme chiffré et lève `EncryptedFileError` sauf si un mot de passe est fourni pour le déchiffrement ; un `.ppt` protégé par mot de passe lève `EncryptedPptError`, car le chiffrement RC4 historique n'est pas déchiffré. Une entrée qui n'est ni un ZIP ni un fichier composé OLE est rejetée avec une erreur explicite. Les archives surdimensionnées lèvent `ZipBombError` avant toute analyse.
2. **L'analyse XML** utilise `fast-xml-parser` ; chaque partie devient un arbre d'objets simples (le type `XmlObject`) avec des attributs préfixés par `@_`.
3. **La résolution du thème** charge pour chaque masque son thème, sa table de couleurs (`p:clrMap`), son jeu de polices et son jeu de formats, afin que l'analyse des diapositives puisse résoudre les couleurs de schéma et les références de style (voir ci-dessous).
4. **L'analyse des éléments** transforme l'arbre de formes de chaque diapositive en [union discriminée `PptxElement`](/fr/guide/data-model#l-union-pptxelement), convertit les coordonnées EMU en pixels et préserve le XML brut des constructions que le modèle typé ne couvre pas.
5. **`PptxLoadDataBuilder`** assemble tout le reste : sections, diaporamas personnalisés, polices incorporées, masques de notes et de documents, étiquettes, auteurs de commentaires, propriétés du document, la miniature et les avertissements de compatibilité.

### Le pipeline de sauvegarde

`handler.save(slides, options?)` inverse le processus (`PptxHandlerRuntimeSavePipeline.ts`) :

```
PptxSlide[]
  │  resolve conformance class ('preserve' | 'strict' | 'transitional')
  ▼
Reconcile presentation slide list      order, additions, deletions, rels
  ▼
Serialize each slide to OpenXML        elements -> <p:spTree>, embed new media
  ▼
Rebuild [Content_Types].xml            slide overrides + media defaults
  ▼
Comments, masters, layouts             typed mutations applied; untouched
  │                                    parts pass through verbatim
  ▼
Optional Strict conversion             remap namespaces if target is strict
  ▼
JSZip -> Uint8Array                    a valid .pptx / .ppsx / .pptm
```

Deux propriétés de ce pipeline comptent pour la fidélité :

- **Transmission directe par défaut.** Les parties que vous n'avez pas modifiées (masques, mises en page, notes, extensions inconnues, balisage propriétaire) sont reprises de l'archive d'origine octet pour octet ou réémises à partir de leurs arbres d'analyse préservés. La sauvegarde est une réécriture ciblée, pas une régénération complète.
- **Sortie consciente de la conformité.** Le package sauvegardé correspond par défaut à la classe de conformité OOXML du fichier chargé, et peut être forcé en Strict ou Transitional. Voir [Conformité OpenXML](/architecture/openxml-conformance) pour savoir exactement ce qui est remappé.

### Chaîne de résolution des thèmes

Les styles PowerPoint se résolvent via une chaîne d'héritage, que le moteur reproduit fidèlement :

```
Element  ->  Placeholder  ->  Layout  ->  Master  ->  Theme
```

- Les propriétés explicites d'un élément l'emportent toujours.
- Un élément de type espace réservé (un titre, un corps, un pied de page, ...) hérite des valeurs par défaut de texte et de position de l'espace réservé correspondant de sa **mise en page**, qui hérite à son tour du **masque** (`p:txStyles` pour les valeurs par défaut du texte de titre/corps/autre).
- Les références de style (`a:fillRef`, `a:lnRef`, `a:effectRef`, `a:fontRef`) se résolvent par index dans le jeu de formats du **thème**, la couleur de la référence remplaçant la couleur d'espace réservé du thème.
- Les couleurs de schéma (`accent1`, `bg1`, `tx1`, ...) se résolvent via le `p:clrMap` du masque (et tout `p:clrMapOvr` de mise en page) dans le jeu de couleurs du thème.

La chaîne est implémentée par des mixins de runtime dédiés : `PptxHandlerRuntimeThemeLoading`, `...ThemeProcessing`, `...ThemeFormatScheme`, `...ThemeRefResolution`, `...ThemeOverrides`, et le trio d'espaces réservés `...PlaceholderLookup`, `...PlaceholderDefaults`, `...PlaceholderStyles`. Plusieurs masques sont pris en charge, chacun avec sa propre table de couleurs et son propre jeu de formats.

### Moteur de géométrie

`src/core/pptx/core/geometry/` (ainsi que les tables de formes prédéfinies partagées de `src/core/geometry/`) transforme la géométrie DrawingML en tracés affichables :

- **Formes prédéfinies** : définitions du catalogue de formes prédéfinies ECMA-376, regroupées par famille (`preset-shape-definitions-arrows.ts`, `-flowchart.ts`, `-action-buttons.ts`, `-callouts`, ...), chacune exprimée avec les formules de guides de la spécification.
- **Évaluation des formules de guides** (`guide-formula-eval.ts` et fichiers associés) : implémente le langage de formules ECMA-376 (`*/`, `+-`, `pin`, `at2`, `cos`, ...) pour que la géométrie des formes réagisse correctement aux valeurs d'ajustement, les poignées en losange jaune que l'on peut faire glisser dans PowerPoint.
- **Géométrie personnalisée** (`custom-geometry-parser.ts`, `freeform-builder.ts`) : analyse les commandes de tracé `a:custGeom` et construit des formes libres.
- **Tracés de détourage** (`preset-clip-paths-core.ts`, `-extended.ts`) : produisent des tracés de détourage CSS/SVG afin que le contenu HTML (images, texte) puisse être découpé selon n'importe quelle forme prédéfinie.
- **Géométrie des connecteurs** (`connector-geometry.ts`) : connecteurs coudés et courbes, pointes de flèche et gestion des retournements. (Le **routage** dynamique des connecteurs autour des obstacles vit dans le routeur A\* de la couche partagée.)

### Convertisseur

`src/core/pptx/converter/` implémente la conversion PPTX vers Markdown avec un patron de registre : chaque type d'élément a un processeur (`shape-element-processor`, `table-element-processor`, `ole-element-processor`, ...) enregistré sur son discriminant `type`, et `PptxMarkdownConverter` répartit par élément. Le même répertoire héberge l'exportateur SVG et le convertisseur OMML vers LaTeX utilisé pour les équations.

## Décisions de conception clés

| Décision                                | Justification                                                                                                                                                                                                                                                               |
| --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Rendu basé sur CSS** (pas Canvas)     | Texte net à tout zoom, accessibilité native, interactivité DOM et style CSS standard.                                                                                                                                                                                       |
| **Moteur composé par mixins**           | De nombreux petits modules de runtime à préoccupation unique gardent chaque capacité isolée et testable. Les nouvelles capacités (y compris les nouveaux types d'éléments) s'ajoutent comme de nouveaux mixins, pas comme des fichiers plus gros.                           |
| **Union discriminée pour les éléments** | TypeScript restreint au bon type d'élément via le champ `type`, ce qui donne une sûreté de typage complète sans conversion.                                                                                                                                                 |
| **Chaîne de résolution des thèmes**     | Élément, espace réservé, mise en page, masque, thème reproduit l'héritage de style propre à PowerPoint.                                                                                                                                                                     |
| **Unités EMU en interne**               | PowerPoint utilise les English Metric Units (914 400 EMU par pouce ; 9 525 EMU par pixel à 96 DPI). Les éléments analysés exposent des valeurs en pixels pour faciliter les calculs de mise en page ; les valeurs EMU exactes sont conservées là où l'aller-retour l'exige. |
| **Sauvegarde par transmission directe** | Seules les parties modifiées sont réécrites ; tout le reste fait l'aller-retour à l'identique, de sorte que le balisage inconnu et les extensions propriétaires survivent.                                                                                                  |
| **Logique partagée, bindings fins**     | Toute la logique de visionneuse indépendante du framework vit une seule fois dans `ooxml-ui/pptx` ; les bindings ne sont que des couches de vue, ce qui garde les cinq frameworks à parité.                                                                                 |

## Lectures connexes

- [Le modèle PptxData](/fr/guide/data-model) - la forme complète d'une présentation analysée, les types d'éléments et les unités.
- [Conformité OpenXML](/architecture/openxml-conformance) - gestion de Strict et Transitional et contrat de conformité.
- [Aperçu du package Core](/core/) - la référence de l'API publique.
