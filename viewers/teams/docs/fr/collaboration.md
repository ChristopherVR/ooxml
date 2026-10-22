# Collaboration

La collaboration est ce qu’est OpenTeams : chaque canal, message, réaction et signal de présence est un état partagé que plusieurs personnes modifient en même temps. Cette page explique comment il est construit, ce que montre la démo en direct et ce qu’elle ne montre pas.

## Comment ça marche

- **Le document partagé est un CRDT Yjs.** Les canaux vivent dans un seul `Y.Map`, les messages dans un `Y.Map` par canal et les réactions dans une table à plat indexée par message, émoji et utilisateur, de sorte que les modifications et réactions concurrentes fusionnent au lieu de s’écraser mutuellement. La logique se trouve dans `ooxml-core/teams`, construite sur le domaine `ooxml-core/collab`, indépendant des formats.
- **La présence est l’awareness de Yjs** : qui est en ligne, qui est en train d’écrire et une disponibilité que vous définissez (disponible, occupé, absent). Elle est indicative et non authentifiée, comme toute awareness Yjs.
- **Les appels utilisent WebRTC.** La signalisation passe par votre serveur (ou par `BroadcastChannel` en mode local) ; le flux média circule de pair à pair en maillage complet, de sorte que les appels conviennent à une poignée de personnes.
- **Le transport est interchangeable.** Il existe deux modes : `server` (votre serveur parle le protocole de synchronisation y-websocket, relaie la signalisation des appels et stocke éventuellement les fichiers) et `local` (les onglets d’un même navigateur partagent tout via `BroadcastChannel`).

## Démo en direct : deux personnes, une salle, n’importe quel framework

Les deux volets ci-dessous sont deux applications de démo distinctes dans une même salle locale. Ada exécute le `<teams-app>` vanilla et Bob exécute l’intégration React, et pourtant ils partagent un même espace de travail : une salle n’est qu’un nom (`?room=`), et chaque intégration parle le même protocole du cœur. Elles s’exécutent en mode local, donc elles ne communiquent que par votre navigateur. Écrivez un message en tant qu’Ada et regardez-le apparaître chez Bob.

Sur la [démo en direct](/fr/#live-demo) de la page d’accueil, chaque volet a son propre sélecteur de framework, ce qui vous permet d’associer deux d’entre Vanilla, React, Vue, Angular, Svelte et Solid.

<iframe
	src="/ooxml/teams/demo/?name=Ada&room=docs-collab"
	title="Démo OpenTeams en tant qu’Ada"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/teams/demo-react/?name=Bob&room=docs-collab"
	title="Démo React OpenTeams en tant que Bob"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

Cela fonctionne de même entre de vrais onglets, quelle que soit l’association : ouvrez [la démo vanilla en tant qu’Ada](/demo/?name=Ada&room=docs-collab){target="_self"} et, par exemple, [la démo Vue en tant que Bob](/demo-vue/?name=Bob&room=docs-collab){target="_self"}. Toutes les démos acceptent les mêmes paramètres `name` et `room` ; en mode local, seuls les cadres et les onglets d’un même navigateur peuvent se rencontrer, donc pour des personnes sur des machines différentes, vous faites pointer chaque intégration vers votre propre serveur (voir plus bas).

## Ce qui est pris en charge et ce qui ne l’est pas

| Pris en charge                                                             | Non pris en charge                                                                                                          |
| -------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| Discussions, réponses et réactions concurrentes fusionnées par un CRDT Yjs | Chiffrement de bout en bout : votre serveur peut tout lire et tout stocker                                                  |
| Présence et indicateurs de saisie                                          | Contrôle d’accès par équipe et SSO ; l’identité est une déclaration sauf si vous activez les comptes facultatifs du serveur |
| Collaboration dans un même navigateur sans serveur (mode local, ci-dessus) | Collaboration entre machines différentes sans un serveur que vous exploitez                                                 |
| Collaboration entre machines via votre propre serveur de synchronisation   | Un service hébergé : aucun n’est fourni, et les démos Pages n’ont aucun serveur derrière                                    |
| Appels avec audio, vidéo, partage d’écran et main levée (maillage, ~12)    | Appels de la taille d’un webinaire, enregistrement, notifications                                                           |

## Entre machines

Exécutez le serveur de référence (`npx openteams-server`) ou tout serveur qui respecte les trois contrats (synchronisation, signalisation, fichiers facultatifs) et définissez `config.mode` sur `'server'`. Voir [Apportez votre propre serveur](/fr/server) et [Déploiement](/fr/deploy). La co-édition de fichiers Office ne fait pas partie d’OpenTeams : les fichiers sont partagés sous forme de liens, et un hôte peut les ouvrir dans la visionneuse Word, Excel, PowerPoint ou Visio, chacune avec sa propre approche de la collaboration.
