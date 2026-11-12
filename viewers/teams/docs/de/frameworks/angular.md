# Angular

`openteams-angular-viewer` (Peer: `@angular/core` ab Version 17) stellt eine eigenständige Komponente
`<teams-workspace>` und einen injizierbaren `TeamsService` bereit.

```bash
npm install openteams-angular-viewer
```

Das Paket ist ein gebündeltes ES-Modul, dessen Decorators zur Laufzeit angewendet werden. Die verwendende App
braucht daher den JIT-Compiler von Angular.

## Die Komponente {#the-component}

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

Eingaben: `workspaceId`, `userName`, `userId`, `config`, `uploadFile`, `openers`, `embeds`, `locale`. Ausgaben: `ready`,
`openFile` (`{ detail, event }`), `configChange`.

## Der Service {#the-service}

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

`client` und `state` sind Angular-Signale. `disconnect()` (wird auch beim Zerstören aufgerufen) gibt den Client frei.
