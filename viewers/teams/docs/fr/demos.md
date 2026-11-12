# Les démos en direct

Une démo par liaison est publiée avec ce site. Toutes font tourner le même espace de travail et
acceptent les mêmes paramètres de requête (`?name=Ada&room=acme`).

| Démo                                                                                         | Ce qu’elle montre                                                                                  | Source                |
| -------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- | --------------------- |
| [Vanilla](/demo/){target="_self"} (aussi à [/demo-vanilla/](/demo-vanilla/){target="_self"}) | L’application complète : un `<teams-app>` monté par `mountTeams`.                                  | `demos/teams/vanilla` |
| [React](/demo-react/){target="_self"}                                                        | `<Teams />` à gauche et un panneau écrit à la main, s’appuyant sur le hook `useTeams()`, à droite. | `demos/teams/react`   |
| [Vue](/demo-vue/){target="_self"}                                                            | Le composant `<Teams>` de `openteams-vue-viewer`.                                                  | `demos/teams/vue`     |
| [Angular](/demo-angular/){target="_self"}                                                    | `<teams-workspace>` de `openteams-angular-viewer`.                                                 | `demos/teams/angular` |
| [Svelte](/demo-svelte/){target="_self"}                                                      | `Teams.svelte` de `openteams-svelte-viewer`.                                                       | `demos/teams/svelte`  |
| [Solid](/demo-solid/){target="_self"}                                                        | `Teams()` de `openteams-solid-viewer`.                                                             | `demos/teams/solid`   |

Comme elles partagent une même origine et un même nom de salle locale, une démo d’un framework peut
communiquer avec une démo d’un autre : ouvrez `/demo-vue/?name=Ada&room=x` et
`/demo-svelte/?name=Bob&room=x` dans deux onglets.

## Elles fonctionnent sans serveur {#they-run-without-a-server}

GitHub Pages ne sert que des fichiers statiques : il n’y a donc aucun serveur derrière ces pages, et
rien n’est simulé pour y ressembler. Les démos sont compilées avec `VITE_TEAMS_STATIC=1`, ce qui les
lance dans le **mode local** du cœur :

- la discussion, la présence et la signalisation des appels circulent entre les onglets de **votre
  seul navigateur**, via `BroadcastChannel` ;
- le document partagé est conservé dans le `localStorage` de votre navigateur, il survit donc à un
  rechargement ;
- les fichiers ne sont partagés que par leur nom, car il n’existe nulle part où les envoyer ;
- les appels relient les onglets du même navigateur (votre caméra et votre micro, avec votre
  autorisation).

Chaque démo affiche un avis qui le précise, accompagné d’un lien qui ouvre un second onglet comme
une autre personne (le paramètre `?name=`). Écrivez dans un onglet, et le message apparaît dans
l’autre.

Si vous avez enregistré votre propre serveur dans les paramètres **Serveur** de la démo vanilla, il
est utilisé à la place ; effacez-le à cet endroit (ou effacez le stockage du site) pour revenir au
mode local.

## Utiliser votre propre serveur {#use-your-own-server}

Une utilisation réelle nécessite un serveur : lancez `npx openteams-server --insecure-dev` (pour le
développement local uniquement) et consultez [Apportez votre serveur](/fr/server). Depuis les sources,
`bun run dev` démarre le serveur de référence et les deux démos qui s’y connectent (voir
[Premiers pas](/fr/getting-started#run-it-from-source)).
