---
title: Tester l'intégration OneDrive
description: Vérifier la connexion OneDrive Office partagée avec des transports de test locaux et une acceptation avec un compte Microsoft autorisé.
---

# Tester l'intégration OneDrive

La connexion OneDrive est partagée par Word, Excel, PowerPoint et Visio. Ses transports automatisés n'appartiennent qu'aux fixtures de test. Il n'existe ni interface de simulation livrée, ni service factice public, ni démonstration de simulation OneDrive à configurer ou déployer.

Utilisez le [guide de test Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) pour les commandes, les vérifications par format et la matrice d'acceptation avec un vrai compte. Le [guide d'intégration](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) décrit le câblage des packages et des frameworks, l'authentification, le partage et la récupération après conflit.

Les tests de régression automatisés peuvent vérifier la séquence de requêtes Graph et le câblage d'enregistrement de l'éditeur sans identifiants Microsoft. Ils couvrent les instantanés et l'identité des fichiers, les restrictions d'autorisation, les conditions de révision, la conservation des modifications locales, les copies cloud distinctes et la gestion des erreurs. Ils ne peuvent pas prouver la connexion Microsoft, le consentement, l'accès conditionnel, le CORS de l'origine déployée ni la fidélité de PowerPoint natif.

Microsoft Graph Explorer permet des vérifications manuelles de l'API avec l'enregistrement d'application existant de Microsoft. Il ne valide ni l'enregistrement propre à l'application ni son flux de connexion. Utilisez une présentation jetable et suivez le guide de test ; ne livrez pas l'identifiant client ni les identifiants empruntés de Graph Explorer comme solution d'authentification de l'application.

Pour une modification réelle via la suite, configurez le fournisseur de jetons Microsoft de l'hôte avant le chargement de son point d'entrée, puis utilisez **Ouvrir depuis OneDrive**. Les comptes réels personnels ou professionnels et les règles des locataires SharePoint restent des cas d'acceptation obligatoires. Aucun test sans enregistrement n'est une preuve de coédition Microsoft en direct ni d'édition simultanée garantie sans conflit.
