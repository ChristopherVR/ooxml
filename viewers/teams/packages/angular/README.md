# openteams-angular-viewer

Angular binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you a standalone component **and** a raw service with signals.

[Source](https://github.com/ChristopherVR/teams-viewer) | [Server](https://www.npmjs.com/package/openteams-server)

## Install

```bash
npm install openteams-angular-viewer
```

`@angular/core` (17 or later) is a peer dependency. The `<teams-app>` element is bundled in;
`ooxml-core`, `ooxml-ui` and `lit` are installed as regular dependencies. The package is plain
ES2022 JavaScript with decorators applied at runtime (no Ivy partial compilation), so it needs the
JIT compiler (`@angular/compiler`) available, as with any non-AOT library component.

## Use

```ts
import { Component } from '@angular/core';
import { TeamsWorkspaceComponent, TeamsService } from 'openteams-angular-viewer';

@Component({
	selector: 'app-root',
	standalone: true,
	imports: [TeamsWorkspaceComponent],
	providers: [TeamsService],
	template: `<teams-workspace workspaceId="acme" userName="Ada" [config]="config" />`,
})
export class AppComponent {
	config = { mode: 'server' as const, syncUrl: 'wss://teams.example.com/sync', signalingUrl: 'wss://teams.example.com/signal' };
}
```

For your own UI, call `TeamsService.connect({ workspaceId, user, config })` and read
`state()` in templates. Without `config` the workspace runs in local mode (tabs of one browser).

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: your server can read chat. Mesh calls scale to about a dozen people, and
authorization is the server's job.

Licence: Apache-2.0.
