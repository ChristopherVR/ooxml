---
title: Aperçu de la documentation
description: Une carte de la documentation pptx-viewer - quelle section lire selon que vous intégrez une visionneuse, manipulez des fichiers .pptx par code, utilisez l'éditeur ou automatisez avec des agents.
---

# Aperçu de la documentation

La documentation est organisée par objectif. Utilisez cette page pour trouver la bonne section, puis suivez l'ordre de lecture proposé.

## Choisissez un parcours

| Vous voulez...                                                                                 | Commencez par                                                                                                                                                                                                              | Ensuite                                                                                                            |
| ---------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| Intégrer une visionneuse ou un éditeur PowerPoint dans une application web                     | La page de démarrage de votre framework : [React](/react/getting-started), [Vue 3](/vue/getting-started), [Angular](/angular/getting-started), [Svelte 5](/svelte/getting-started), [Vanilla JS](/vanilla/getting-started) | La référence props/API de la même section, puis [Thèmes](/guide/theming) et [Localisation](/fr/guide/localization) |
| Lire, modifier ou générer des fichiers `.pptx` par code                                        | [Core : chargement](/core/loading) et [l'API Builder](/core/builder)                                                                                                                                                       | [Édition](/core/editing), [Enregistrement](/core/saving), [le modèle de données](/fr/guide/data-model)             |
| Convertir des présentations en Markdown, images, PDF ou SVG                                    | [Convertisseur Markdown](/core/converter) et [Export SVG](/core/svg-export)                                                                                                                                                | La page Export de votre binding pour les formats navigateur                                                        |
| Masquer, verrouiller ou réaffecter des parties de l'interface de la visionneuse depuis le code | [Personnalisation de l'interface](/guide/customization) (en anglais)                                                                                                                                                       | Recettes sur cette page : mode kiosque, ruban allégé, paramètres épinglés                                          |
| Utiliser l'éditeur (sans écrire de code)                                                       | [Guide utilisateur](/fr/user/)                                                                                                                                                                                             | [Édition des diapositives](/fr/user/editing), [Raccourcis clavier](/fr/user/shortcuts)                             |
| Laisser des agents IA ou des scripts travailler sur des présentations                          | [MCP et outils](/packages/mcp)                                                                                                                                                                                             | [CLI du core](/core/cli)                                                                                           |
| Comprendre le fonctionnement interne de la bibliothèque                                        | [Architecture](/fr/guide/architecture)                                                                                                                                                                                     | [Conformité OOXML](/architecture/openxml-conformance), [Limitations](/fr/guide/limitations)                        |

## Nouveau sur le projet

Si aucun parcours ci-dessus ne correspond encore, lisez ces trois pages dans l'ordre :

1. [Qu'est-ce que pptx-viewer ?](/fr/guide/introduction) : la famille de packages et le rôle de chacun.
2. [Installation](/fr/guide/installation) : quel package installer pour votre stack.
3. [Démarrage rapide](/fr/guide/quick-start) : quatre flux complets avec l'API publique.

## Les sections en un coup d'œil

| Section                                                                                                                | Public                          | Contenu                                                                                             |
| ---------------------------------------------------------------------------------------------------------------------- | ------------------------------- | --------------------------------------------------------------------------------------------------- |
| **Guide du développeur** (cette section)                                                                               | Développeurs                    | Concepts communs à tous les packages : architecture, modèle de données, thèmes, i18n, limitations.  |
| **[Core](/core/)**                                                                                                     | Développeurs                    | Le moteur headless : charger, éditer, construire, sauvegarder, convertir, chiffrer. Sans interface. |
| **[React](/react/)** / **[Vue](/vue/)** / **[Angular](/angular/)** / **[Svelte](/svelte/)** / **[Vanilla](/vanilla/)** | Développeurs                    | Une section par binding UI : démarrage, props/API, thèmes, export, collaboration.                   |
| **[Guide utilisateur](/fr/user/)**                                                                                     | Utilisateurs de l'éditeur       | Comment afficher, éditer, présenter, exporter et collaborer dans l'éditeur. Sans code.              |
| **[MCP et outils](/packages/mcp)**                                                                                     | Auteurs d'agents/automatisation | Le serveur MCP, ses 73 outils et la CLI.                                                            |
| **[Versions](/releases/)**                                                                                             | Tout le monde                   | Notes de version par package.                                                                       |
