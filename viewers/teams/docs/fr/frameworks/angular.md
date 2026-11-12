# Angular

`openteams-angular-viewer` (dépendance homologue : `@angular/core` 17 ou plus récent) vous fournit
un composant autonome `<teams-workspace>` et un service injectable `TeamsService`.

```bash
npm install openteams-angular-viewer
```

Le paquet est un module ES regroupé, dont les décorateurs sont appliqués à l’exécution ; l’application
qui l’utilise a donc besoin du compilateur JIT d’Angular.

## Le composant {#the-component}

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

Entrées : `workspaceId`, `userName`, `userId`, `config`, `uploadFile`, `openers`, `embeds`, `locale`.
Sorties : `ready`, `openFile` (`{ detail, event }`), `configChange`.

## Le service {#the-service}

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

`client` et `state` sont des signaux Angular ; `disconnect()` (aussi appelé à la destruction)
libère le client.
