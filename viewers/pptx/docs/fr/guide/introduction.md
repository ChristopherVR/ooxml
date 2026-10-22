---
title: Qu'est-ce que pptx-viewer ?
description: Un aperçu général du monorepo TypeScript pptx-viewer pour analyser, éditer, afficher et convertir des fichiers PowerPoint avec React, Vue 3, Angular, Svelte 5 et JavaScript natif.
---

# Qu'est-ce que pptx-viewer ?

`pptx-viewer` est un monorepo TypeScript pour **analyser, éditer, afficher et convertir** des fichiers Microsoft PowerPoint (`.pptx`), dans le navigateur et dans Node.js. Il fonctionne entièrement en mémoire sur l'archive ZIP OpenXML, sans aucune dépendance native.

`pptx-viewer` couvre le cycle complet : charger une présentation existante, modifier son modèle de données structuré, l'afficher et la sauvegarder dans un fichier `.pptx` valide. Le même moteur principal alimente des composants pour **React**, **Vue 3**, **Angular** et **Svelte 5**, ainsi qu'une version **JavaScript natif** pour les projets sans framework.

## Ce qu'il fait

Le SDK fournit neuf capacités principales :

1. **Analyser** les fichiers `.pptx` (ainsi que `.ppsx` / `.pptm` / `.potx`, et le format binaire historique `.ppt` de PowerPoint 97-2003) depuis un `ArrayBuffer` brut en un modèle [`PptxData`](/fr/guide/data-model) structuré.
2. **Créer** des présentations de toutes pièces avec une API de construction fluide.
3. **Afficher** les diapositives en tant que composants React, Vue, Angular, Svelte ou JavaScript natif interactifs.
4. **Éditer** les présentations par programmation ou via l'éditeur WYSIWYG intégré.
5. **Sauvegarder** les modifications dans un fichier `.pptx` valide (compatible aller-retour).
6. **Convertir** les présentations en Markdown avec extraction optionnelle des médias.
7. **Exporter** les diapositives en images (PNG/JPEG), SVG, PDF, GIF ou vidéo.
8. **Collaborer** en temps réel via Yjs CRDT avec suivi de présence.
9. **Chiffrer/Déchiffrer** les fichiers PPTX protégés par mot de passe (AES-128/256).

Le moteur gère les éléments de présentation, des centaines de formes prédéfinies `ST_ShapeType`, les graphiques, les styles de tableau intégrés, SmartArt, les modèles 3D, les animations et transitions (y compris morphe), les thèmes, les masques de diapositives, les médias incorporés, les métafichiers EMF/WMF, les objets OLE, l'encre numérique, les signatures numériques, le chiffrement, la préservation des macros VBA et la conformité OOXML Strict.

## Pour qui est-il conçu ?

- **Équipes produit qui intègrent une visionneuse de documents** : affichez les présentations envoyées par les utilisateurs dans une application web sans envoyer les fichiers à un service de conversion.
- **Équipes qui construisent un éditeur** : les bindings incluent une surface d'édition WYSIWYG complète (ruban, inspecteur, annuler/rétablir, collaboration) que vous pouvez activer fonctionnalité par fonctionnalité.
- **Backend et automatisation** : le package Core est indépendant du framework et fonctionne dans Node.js ; vous pouvez donc générer, inspecter, comparer ou convertir des présentations dans des scripts, des serveurs et la CI.
- **Pipelines IA et agents** : `pptx-viewer-mcp` expose le moteur sous forme d'appels d'outils MCP, et le convertisseur Markdown transforme les présentations en texte adapté aux LLM.

## Modèle de rendu

`pptx-viewer` ne rastérise pas les diapositives, ni côté serveur ni sur un `<canvas>`. Les diapositives sont rendues en **HTML, CSS et SVG dans le DOM** :

- Le texte est du vrai texte HTML : sélectionnable, consultable, traduisible et lisible par les lecteurs d'écran.
- Les formes et connecteurs sont du SVG avec des tracés de détourage calculés par le moteur de géométrie.
- Les tableaux sont des éléments HTML `<table>` ; les graphiques sont du SVG en ligne construit à partir des données de graphique analysées.
- Le zoom est une transformation CSS : les diapositives restent nettes à toute échelle et sur toute densité d'affichage.
- Comme chaque élément est un nœud du DOM, l'éditeur bénéficie gratuitement de la sélection, des poignées de déplacement, de l'édition de texte en ligne et de l'accessibilité.

La rastérisation n'a lieu que sur demande : l'export raster (PNG/JPEG/PDF/GIF/vidéo) tente d'abord une capture SVG `foreignObject` (le moteur du navigateur peint lui-même le DOM, en préservant `backdrop-filter`, les propriétés personnalisées et les transformations 3D), se rabat sur la re-rastérisation du pipeline d'export SVG vectoriel existant, et ne dessine le DOM via `html2canvas-pro` qu'en dernier recours.

::: tip Compromis
Le CSS ne peut pas exprimer pixel pour pixel tous les effets PowerPoint, et le chemin de repli `html2canvas-pro` (utilisé lorsqu'une ressource ne peut pas être incorporée pour la capture SVG) a ses propres lacunes, par exemple l'approximation des modes de fusion par l'opacité. Voir [Limitations](/fr/guide/limitations) pour la liste complète.
:::

## Les packages

Le monorepo publie huit packages.

| Package                           | Nom npm                      | Objectif                                                                                                      |
| --------------------------------- | ---------------------------- | ------------------------------------------------------------------------------------------------------------- |
| [**Core**](/core/)                | `pptx-viewer-core`           | Analyser, créer, éditer, sérialiser et convertir des fichiers PPTX. Indépendant du framework.                 |
| [**React**](/react/)              | `pptx-react-viewer`          | Visionneuse, éditeur et présentateur React avec barre d'outils, inspecteur, collaboration et export.          |
| [**Vue 3**](/vue/)                | `pptx-vue-viewer`            | Visionneuse/éditeur Vue 3 construit sur le même moteur, avec le même ensemble de fonctionnalités.             |
| [**Angular**](/angular/)          | `pptx-angular-viewer`        | Visionneuse/éditeur Angular construit sur le même moteur, avec le même ensemble de fonctionnalités.           |
| [**Vanilla JS**](/vanilla/)       | `pptx-vanilla-viewer`        | Binding sans framework construit sur le même moteur : DOM natif, une fonction de fabrique, aucune dépendance. |
| [**Svelte**](/svelte/)            | `pptx-svelte-viewer`         | Composant Svelte 5 construit sur le même moteur, avec le même ensemble de fonctionnalités.                    |
| [**Outils / MCP**](/packages/mcp) | `pptx-viewer-mcp`            | 73 fonctions d'outils PPTX, un serveur MCP pour les agents IA et le codec de collaboration Y.Doc.             |
| **Installateur**                  | `@christophervr/pptx-viewer` | CLI interactif qui intègre le bon package de visionneuse dans votre projet.                                   |

### Comment les éléments s'articulent

Les cinq packages de bindings UI s'appuient sur une couche de rendu partagée, qui s'appuie à son tour sur Core :

```
pptx-react-viewer   ┐
pptx-vue-viewer     │
pptx-angular-viewer ├── ooxml-ui/pptx ── ooxml-core/pptx
pptx-vanilla-viewer │                               ├── emf-converter
pptx-svelte-viewer  ┘                               └── mtx-decompressor
```

- **`pptx-viewer-core`** possède tout ce qui concerne le format de fichier : les pipelines de chargement et de sauvegarde, le modèle de données typé, la résolution des thèmes, le moteur de géométrie, le chiffrement et les convertisseurs Markdown/SVG. Il n'a pas d'interface et fonctionne partout où JavaScript fonctionne.
- **`ooxml-ui/pptx`** publie le rendu partagé, le style, les composants web, le chargement dans le navigateur et les vues d'export depuis `src/ui/src/pptx`. Chaque binding importe cette API publique. Les opérations d'édition et de collaboration sans DOM vivent dans `ooxml-core/pptx/editor` ; le package privé `pptx-viewer-shared` n'est plus qu'une façade de compatibilité.
- **Les bindings** sont de fines couches de vue. Chacun traduit les mêmes données de rendu partagées dans l'idiome de son framework (JSX, modèles SFC, modèles Angular, runes Svelte ou DOM natif), ce qui explique que les cinq bindings s'affichent de façon identique et restent à parité fonctionnelle.
- **`emf-converter`** (métafichier EMF/WMF vers PNG) et **`mtx-decompressor`** (polices incorporées MicroType Express) sont des dépendances npm autonomes de Core.
- **`pptx-viewer-mcp`** s'appuie sur `pptx-viewer-core` pour exposer des surfaces d'appel d'outils et de collaboration aux agents IA.

## Exemples minimaux

::: code-group

```ts [Load, edit, save]
import { PptxHandler } from 'pptx-viewer-core';

const handler = new PptxHandler();
const data = await handler.load(arrayBuffer);

// data.slides is a typed, mutable model
data.slides[0].elements.filter((el) => el.type === 'text').forEach((el) => console.log(el.text));

const bytes = await handler.save(data.slides); // => Uint8Array (.pptx)
```

```ts [Create from scratch]
import { PptxHandler } from 'pptx-viewer-core';

const { handler, data, createSlide } = await PptxHandler.create({
	title: 'Q4 Report',
	theme: { colors: { accent1: '#FF6B6B' } },
});

data.slides.push(createSlide('Blank').addText('Hello', { fontSize: 36 }).build());

const bytes = await handler.save(data.slides);
```

:::

## Étapes suivantes

- [Installation](/fr/guide/installation) - installer les packages et configurer le développement local.
- [Démarrage rapide](/fr/guide/quick-start) - exemples d'utilisation de bout en bout.
- [Architecture](/fr/guide/architecture) - fonctionnement des pipelines de chargement/sauvegarde et de la conception en couches.
- [Le modèle PptxData](/fr/guide/data-model) - le modèle typé que vous éditez.
- [Aperçu du package Core](/core/) - le moteur d'analyse, d'édition et de sérialisation.
- [Aperçu du package React](/react/) - le composant visionneuse/éditeur.
- [Limitations](/fr/guide/limitations) - mises en garde importantes à lire avant d'adopter.
