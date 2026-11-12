# Angular

`openteams-angular-viewer` (peer: `@angular/core` 17 o posterior) te ofrece un componente standalone
`<teams-workspace>` y un `TeamsService` inyectable.

```bash
npm install openteams-angular-viewer
```

El paquete es un módulo ES empaquetado con sus decoradores aplicados en tiempo de ejecución, así que
la aplicación que lo consume necesita el compilador JIT de Angular.

## El componente {#the-component}

```ts
import { Component } from '@angular/core';
import { TeamsWorkspaceComponent } from 'openteams-angular-viewer';

@Component({
	standalone: true,
	imports: [TeamsWorkspaceComponent],
	template: `<teams-workspace
		workspaceId="acme"
		userName="Ada"
		[config]="config"
		(openFile)="onOpen($event.detail)"
	/>`,
})
export class WorkspaceComponent {
	config = {
		mode: 'server' as const,
		syncUrl: 'wss://teams.example.com/sync',
		signalingUrl: 'wss://teams.example.com/signal',
		iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	};
	onOpen(detail: { url: string | undefined }) {
		console.log(detail.url);
	}
}
```

Entradas: `workspaceId`, `userName`, `userId`, `config`, `uploadFile`, `openers`, `embeds`, `locale`.
Salidas: `ready`, `openFile` (`{ detail, event }`), `configChange`.

## El servicio {#the-service}

```ts
import { TeamsService } from 'openteams-angular-viewer';

@Component({ providers: [TeamsService], template: `...` })
export class Panel {
	constructor(readonly teams: TeamsService) {
		teams.connect({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
	}
}
// template: teams.state()?.channels, teams.client()?.send({ text: 'hi' })
```

`client` y `state` son señales de Angular; `disconnect()` (también se llama al destruir el componente)
libera el cliente.
