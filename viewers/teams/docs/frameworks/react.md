# React

`openteams-react-viewer` (peer: `react` 18 or later) gives you the full UI as `<Teams />` and the
raw client as hooks.

```bash
npm install openteams-react-viewer react
```

## The component

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

Props are applied after every render and only changed values reach the element. The ref is the
`<teams-app>` element itself (its `client` getter returns the core client).

## The hooks

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

`useTeams` is `useTeamsClient` plus `useTeamsState` (`useSyncExternalStore`). The client is
recreated when the workspace, user or server change and destroyed on unmount; keep `options`
stable (`useMemo`).

The [React demo](/demo-react/){target="_self"} uses both on one page: `demos/react/src/App.tsx`.
