# Solid

`openteams-solid-viewer` (dépendance homologue : `solid-js` 1.9) vous fournit l’élément au moyen
d’une simple fonction, et le client brut sous forme de signaux.

```bash
npm install openteams-solid-viewer solid-js
```

## Le composant {#the-component}

`Teams(props)` renvoie l’élément `<teams-app>`. Il n’a donc pas besoin d’une transformation JSX dans
le paquet ; utilisez-le dans votre JSX :

```tsx
import type { TeamsServerConfig } from 'ooxml-core/teams';
import { Teams } from 'openteams-solid-viewer';

export function Workspace(props: { config: TeamsServerConfig }) {
	return (
		<div style={{ height: '100%' }}>
			{Teams({
				workspaceId: 'acme',
				userName: 'Ada',
				get config() {
					return props.config;
				},
				onOpenFile: (detail) => console.log(detail.url),
			})}
		</div>
	);
}
```

Les propriétés sont réappliquées de façon réactive (utilisez des accesseurs `get` pour les valeurs
réactives), et les gestionnaires sont lus au moment où l’événement se produit.

## La primitive {#the-primitive}

```tsx
import { createTeamsClient } from 'openteams-solid-viewer';

const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user, config }));
// state()?.channels, client()?.send({ text: 'hi' })
```

Le client est recréé lorsque `options()` change, et détruit avec son propriétaire.
