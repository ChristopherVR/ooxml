# Collaboration {#collaboration}

Fichier > Partager démarre une session en direct entre des fenêtres d'un même navigateur. Le partage entre appareils n'est pas encore disponible.

::: warning Partiel, un seul navigateur
Cette version n'inclut ni service hébergé ni transport serveur. Rien ne quitte votre navigateur et aucun document n'est envoyé.
:::

## Ce qui existe aujourd'hui {#what-exists-today}

- Une session en direct fondée sur `ooxml-core/collab` (Yjs) via un `BroadcastChannel`, de sorte que les fenêtres et onglets d'un même navigateur peuvent partager un dessin.
- Tout le paquet VSDX est partagé, et le runtime de collaboration n'est chargé qu'au démarrage du partage.
- Les modifications distantes arrivent sous la forme d'une seule étape d'historique annulable, et les collaborateurs apparaissent grâce à l'élément de présence partagé de `ooxml-ui`.
- Le nom de session compte de 1 à 64 lettres, chiffres, traits d'union ou tirets bas. Une fenêtre qui rejoint une session dont la salle contient déjà un dessin l'adopte ; sinon, elle alimente la salle.

## Démo en direct : deux fenêtres, n'importe quel framework {#live-demo-two-windows-any-framework}

Une session de partage n'est qu'un nom. Les deux volets ci-dessous sont des terrains de jeu distincts, construits avec des liaisons de frameworks différents, qui rejoignent au chargement la session `docs-collab` : la fenêtre A est le terrain de jeu vanilla et ouvre l'exemple, la fenêtre B est le terrain de jeu React et reçoit le dessin depuis la salle. Modifiez le texte d'une forme dans un volet et regardez l'autre suivre ; l'annulation inverse l'étape distante.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer fenêtre A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer fenêtre B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Associer deux frameworks au choix {#pair-any-two-frameworks}

Chaque terrain de jeu accepte `?share=<session>`, qui renseigne Fichier > Partager et démarre la session. Seul le nom de la session compte, donc n'importe quelle association fonctionne : la [démo en direct](/fr/#live-demo) de la page d'accueil propose un sélecteur de framework pour chaque fenêtre, et dans deux onglets vous pouvez ouvrir, par exemple, `/demo-vue/?sample=1&share=my-room` et `/demo-angular/?share=my-room`. Vous pouvez aussi le faire à la main dans deux fenêtres quelconques : Fichier, puis Partager, puis le même nom de session.

### Pourquoi uniquement dans un même navigateur {#why-it-is-same-browser-only}

La session utilise `ooxml-core/collab` (Yjs) via un `BroadcastChannel`. Un `BroadcastChannel` relie les fenêtres, onglets et cadres d'un même profil de navigateur sur une même origine, quel que soit le framework avec lequel ils ont été construits, et rien d'autre. Deux machines différentes ne peuvent pas se rejoindre ainsi ; cela demande un transport serveur fourni par l'application hôte, et cette version n'en inclut aucun.

## Ce qui est pris en charge et ce qui ne l'est pas {#what-is-and-is-not-supported}

| Pris en charge                                                               | Non pris en charge                                                                                               |
| ---------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| Sessions dans un même navigateur via BroadcastChannel, fondées sur Yjs       | Partage entre appareils : nécessite un transport serveur fourni par l'hôte, aucun n'est livré                    |
| Partage du paquet entier, avec une étape annulable par modification distante | Fusion au niveau des champs : les modifications concurrentes ne sont pas fusionnées, le dernier paquet l'emporte |
| Collaborateurs affichés avec l'élément de présence partagé                   | Commentaires, authentification et rôles                                                                          |
| Fichier > Partager comme point d'entrée                                      | Une API d'intégration documentée pour des transports personnalisés                                               |

Les preuves sont consignées dans le [registre des capacités](/parity) sous « Collaboration and comments ».
