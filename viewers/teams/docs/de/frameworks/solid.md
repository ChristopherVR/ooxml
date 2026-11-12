# Solid

`openteams-solid-viewer` (Peer: `solid-js` 1.9) stellt das Element über eine einfache Funktion bereit und den rohen
Client als Signale.

```bash
npm install openteams-solid-viewer solid-js
```

## Die Komponente {#the-component}

`Teams(props)` gibt das Element `<teams-app>` zurück. Deshalb braucht das Paket keinen JSX-Transform. Verwenden Sie
es innerhalb Ihres JSX:

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

Props werden reaktiv erneut angewendet (verwenden Sie Getter für reaktive Werte), und Handler werden gelesen, wenn das
Ereignis auftritt.

## Das Primitiv {#the-primitive}

```tsx
import { createTeamsClient } from 'openteams-solid-viewer';

const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user, config }));
// state()?.channels, client()?.send({ text: 'hi' })
```

Der Client wird neu erzeugt, wenn sich `options()` ändert, und zusammen mit seinem Besitzer zerstört.
