# React

`openteams-react-viewer` (Peer: `react` ab Version 18) stellt die vollständige Oberfläche als `<Teams />` bereit und den
rohen Client als Hooks.

```bash
npm install openteams-react-viewer react
```

## Die Komponente {#the-component}

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

Props werden nach jedem Render angewendet, und nur geänderte Werte erreichen das Element. Die Ref ist das Element
`<teams-app>` selbst (sein Getter `client` liefert den Core-Client).

## Die Hooks {#the-hooks}

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

`useTeams` ist `useTeamsClient` plus `useTeamsState` (`useSyncExternalStore`). Der Client wird neu erzeugt, wenn sich
Arbeitsbereich, Benutzer oder Server ändern, und beim Aushängen zerstört. Halten Sie `options` stabil (`useMemo`).

Die [React-Demo](/demo-react/){target="_self"} nutzt beide auf einer Seite: `demos/teams/react/src/App.tsx`.
