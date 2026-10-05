# Angular

`openteams-angular-viewer` (peer: `@angular/core` 17 or later) gives you a standalone
`<teams-workspace>` component and an injectable `TeamsService`.

```bash
npm install openteams-angular-viewer
```

The package is a bundled ES module with its decorators applied at runtime, so the consuming app
needs the Angular JIT compiler.

## The component

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

Inputs: `workspaceId`, `userName`, `userId`, `config`, `uploadFile`, `openers`. Outputs: `ready`,
`openFile` (`{ detail, event }`), `configChange`.

## The service

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

`client` and `state` are Angular signals; `disconnect()` (also called on destroy) releases the
client.
