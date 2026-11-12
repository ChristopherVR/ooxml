# Solid

`openteams-solid-viewer` (peer: `solid-js` 1.9) te ofrece el elemento mediante una función simple y
el cliente sin procesar como señales.

```bash
npm install openteams-solid-viewer solid-js
```

## El componente {#the-component}

`Teams(props)` devuelve el elemento `<teams-app>`, así que el paquete no necesita transformación de
JSX; úsalo dentro de tu JSX:

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

Las props se vuelven a aplicar de forma reactiva (usa getters para los valores reactivos), y los
manejadores se leen cuando se dispara el evento.

## La primitiva {#the-primitive}

```tsx
import { createTeamsClient } from 'openteams-solid-viewer';

const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user, config }));
// state()?.channels, client()?.send({ text: 'hi' })
```

El cliente se recrea cuando cambia `options()` y se destruye con su propietario.
