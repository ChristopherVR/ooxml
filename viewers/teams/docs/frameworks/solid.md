# Solid

`openteams-solid-viewer` (peer: `solid-js` 1.9) gives you the element through a plain function and
the raw client as signals.

```bash
npm install openteams-solid-viewer solid-js
```

## The component

`Teams(props)` returns the `<teams-app>` element, so it needs no JSX transform in the package; use
it inside your JSX:

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

Props are re-applied reactively (use getters for reactive values) and handlers are read when the
event fires.

## The primitive

```tsx
import { createTeamsClient } from 'openteams-solid-viewer';

const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user, config }));
// state()?.channels, client()?.send({ text: 'hi' })
```

The client is recreated when `options()` changes and destroyed with the owner.
