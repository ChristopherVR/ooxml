# React

`openteams-react-viewer` (peer: `react` 18 o posterior) te ofrece la interfaz completa como
`<Teams />` y el cliente sin procesar como hooks.

```bash
npm install openteams-react-viewer react
```

## El componente {#the-component}

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

Las props se aplican después de cada render y solo llegan al elemento los valores que cambian. La
referencia es el propio elemento `<teams-app>` (su getter `client` devuelve el cliente del núcleo).

## Los hooks {#the-hooks}

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

`useTeams` es `useTeamsClient` más `useTeamsState` (`useSyncExternalStore`). El cliente se recrea
cuando cambian el espacio de trabajo, el usuario o el servidor, y se destruye al desmontar; mantén
`options` estable (`useMemo`).

La [demo de React](/demo-react/){target="_self"} usa ambos en una misma página:
`demos/teams/react/src/App.tsx`.
