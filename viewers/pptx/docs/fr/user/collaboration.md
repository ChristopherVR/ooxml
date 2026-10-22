---
title: Collaboration
description: Démarrez ou rejoignez une session de coédition en direct, lisez les indicateurs de présence, comprenez comment les modifications simultanées fusionnent et diffusez un diaporama.
---

# Collaboration

pptx-viewer prend en charge la coédition en temps réel : plusieurs personnes ouvrent la même présentation et voient les modifications des autres au fur et à mesure.

::: info Optionnel et configuré par le développeur
La collaboration est une fonctionnalité optionnelle. Si vous ne voyez pas de bouton **Partager**, l'application que vous utilisez a été construite sans elle. Développeurs : voir [/react/collaboration](/react/collaboration).
:::

## Démarrer une session

1. Cliquez sur le bouton **Partager** à l'extrémité droite de la rangée d'onglets du ruban (ou **Fichier > Partager > Partager avec des personnes**). La boîte de dialogue **Partager la présentation** s'ouvre sur son onglet **Créer une session**.
2. Renseignez :
   - **Nom de la session** - l'identifiant de la salle (lettres, chiffres, traits d'union, tirets bas). Les collaborateurs ont besoin de ce nom.
   - **Votre nom d'affichage** - la façon dont vous apparaissez aux autres.
   - **Serveur de collaboration** - optionnel. Laissez-le vide pour une session pair-à-pair, ou saisissez l'adresse du serveur de collaboration de votre organisation (une URL `wss://...`).
3. Cliquez sur **Démarrer le partage**.

![Boîte de dialogue de partage/collaboration](/user-guide/collaboration-share-dialog.jpg)

Une fois la session active, la boîte de dialogue affiche un **Lien de partage** avec un bouton **Copier l'URL**. Envoyez ce lien à vos collaborateurs. La boîte de dialogue liste aussi les **Utilisateurs connectés** et la diapositive sur laquelle se trouve chacun. Le bouton Partager de la barre d'outils devient vert et affiche le nombre de participants tant que le partage est actif.

Pour terminer la session, rouvrez la boîte de dialogue et cliquez sur **Arrêter le partage**.

## Rejoindre une session

Ouvrez un lien de partage qu'on vous a envoyé, ou rejoignez manuellement :

1. Cliquez sur **Partager** et passez à l'onglet **Rejoindre une session**.
2. Collez le lien d'invitation, ou saisissez l'identifiant de la salle, dans **Lien d'invitation ou identifiant de salle**.
3. Saisissez votre nom d'affichage et, si la session en utilise un, l'adresse du serveur de collaboration.
4. Cliquez sur **Rejoindre la session**.

## Indicateurs de présence

Pendant une session active, vous verrez :

- **Barre d'avatars** - des cercles colorés dans la barre d'outils pour chaque participant (jusqu'à cinq, puis une pastille de débordement « +N »). Votre propre avatar a une bordure blanche.
- **Curseurs distants** - le pointeur de chaque collaborateur se déplace sur le canevas avec une étiquette de nom colorée. Les curseurs n'apparaissent que pour les personnes qui consultent la même diapositive que vous.
- **Sélections distantes** - lorsque quelqu'un sélectionne un élément, celui-ci reçoit un contour coloré portant son nom, pour que vous sachiez sur quoi il travaille.
- **Barre de suivi** - une bande en haut intitulée **Suivre un collaborateur**. Cliquez sur une personne pour la suivre : votre affichage change alors de diapositive chaque fois qu'elle le fait. Cliquez sur **Arrêter** pour reprendre une navigation indépendante.
- **État de la connexion** - un point dans la barre d'état : **Connecté** en vert (avec le nombre d'utilisateurs), **Connexion...** en jaune, **Déconnecté** en gris, ou **Erreur de connexion** en rouge avec un bouton **Réessayer**.

## Comportement de l'édition simultanée

Les modifications sont fusionnées automatiquement grâce à des structures de données sans conflit :

- Deux personnes qui modifient des diapositives différentes, ou des éléments différents d'une même diapositive, ne s'écrasent jamais mutuellement.
- Deux personnes qui saisissent dans la même zone de texte fusionnent au niveau du caractère, comme dans un éditeur de documents partagé.
- Si votre connexion tombe, vous pouvez continuer à éditer ; vos modifications se synchronisent au retour de la connexion. Surveillez le point d'état : tant qu'il indique **Déconnecté**, les autres ne voient pas encore vos dernières modifications.

::: tip Coordonnez les gros changements
La fusion vous protège des modifications perdues, pas des surprises. Supprimer une diapositive sur laquelle quelqu'un d'autre travaille la supprime tout de même pour tout le monde ; mettez-vous donc d'accord sur qui s'occupe de quelle section pour les restructurations importantes.
:::

## Pair-à-pair ou serveur relais

Le champ **Serveur de collaboration** décide de la façon dont les participants se connectent :

- **Vide (pair-à-pair)** - les onglets du navigateur sur un même appareil se retrouvent toujours, et les personnes sur d'autres appareils se connectent directement à vous via des serveurs de signalisation WebRTC publics. C'est le moyen le plus rapide d'essayer la collaboration ; la démo hébergée fonctionne ainsi.
- **Adresse de serveur renseignée** - tout le monde passe par un serveur de collaboration dédié. Utilisez-le lorsque votre organisation en exploite un : il fonctionne à travers les pare-feu stricts et le NAT là où les connexions directes échouent, garde la salle disponible tant qu'un participant est connecté, et l'application hôte peut y associer une authentification et une persistance côté serveur.

::: warning Utilisez wss:// sur les pages sécurisées
Lorsque l'application est servie en `https`, l'adresse du serveur doit utiliser `wss://` (un WebSocket sécurisé), faute de quoi le navigateur bloquera la connexion.
:::

## Diffuser un diaporama

La diffusion est une collaboration à sens unique : vous présentez, les spectateurs suivent.

1. Cliquez sur **Diaporama > Diffuser** (ou **Présenter en ligne** dans le menu Présenter). La boîte de dialogue **Diffuser le diaporama** s'ouvre.
2. Vérifiez le nom de **Session de diffusion** prérempli et votre **Nom du présentateur**, et renseignez ou videz le **Serveur de collaboration** comme ci-dessus.
3. Cliquez sur **Démarrer la diffusion**. L'application passe en mode présentation et la boîte de dialogue affiche un **Lien spectateur** avec **Copier l'URL**.
4. Partagez le lien spectateur. Les personnes qui l'ouvrent voient vos diapositives avancer en temps réel ; la boîte de dialogue compte vos spectateurs.
5. Cliquez sur **Arrêter la diffusion** lorsque vous avez terminé.

## Les commentaires fonctionnent avec l'édition en direct

Les commentaires (voir [Édition > Commentaires](/fr/user/editing#commentaires)) sont enregistrés avec la présentation et constituent un bon moyen de laisser des retours aux collaborateurs qui ne sont pas en ligne en même temps.

## Suite

- [Édition des diapositives](/fr/user/editing)
- [Raccourcis clavier](/fr/user/shortcuts)
