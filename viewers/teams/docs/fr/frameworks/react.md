# React

`openteams-react-viewer` (dépendance homologue : `react` 18 ou plus récent) vous fournit
l’interface complète sous la forme `<Teams />` et le client brut sous forme de hooks.

```bash
npm install openteams-react-viewer react
```

## Le composant {#the-component}

```tsx
import { Teams } from 'openteams-react-viewer';

export function Workspace({ config }) {
	return (
		<Teams
			workspaceId="acme"
			userName="Ada"
			userId="u-ada"
			config={config}
			onOpenFile={(detail, event) => {
				if (detail.attachment.kind === 'docx') {
					event.preventDefault();
					openInWord(detail.url);
				}
			}}
		/>
	);
}
```

Les propriétés sont appliquées après chaque rendu, et seules les valeurs modifiées atteignent
l’élément. La référence est l’élément `<teams-app>` lui-même (son accesseur `client` renvoie le
client du cœur).

## Les hooks {#the-hooks}

```tsx
import { useTeams, type TeamsClientOptions } from 'openteams-react-viewer';

function Panel({ options }: { options: TeamsClientOptions }) {
	const { client, state } = useTeams(options);
	if (!client || !state) return <p>Connecting…</p>;
	return (
		<ul>
			{state.channels.map((c) => (
				<li key={c.id}>
					<button onClick={() => client.select(c.id)}># {c.name}</button>
				</li>
			))}
		</ul>
	);
}
```

`useTeams` équivaut à `useTeamsClient` plus `useTeamsState` (`useSyncExternalStore`). Le client est
recréé lorsque l’espace de travail, l’utilisateur ou le serveur change, et détruit au démontage ;
gardez `options` stable (`useMemo`).

La [démo React](/demo-react/){target="_self"} utilise les deux sur une même page :
`demos/teams/react/src/App.tsx`.
