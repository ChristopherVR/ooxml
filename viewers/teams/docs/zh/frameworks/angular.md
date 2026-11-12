# Angular

`openteams-angular-viewer`（对等依赖：`@angular/core` 17 或更高版本）提供一个独立的 `<teams-workspace>` 组件和一个可注入的 `TeamsService`。

```bash
npm install openteams-angular-viewer
```

该包是一个已打包的 ES 模块，其装饰器在运行时应用，因此使用它的应用需要 Angular 的 JIT 编译器。

## 组件 {#the-component}

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

输入属性：`workspaceId`、`userName`、`userId`、`config`、`uploadFile`、`openers`、`embeds`、`locale`。输出事件：`ready`、`openFile`（`{ detail, event }`）、`configChange`。

## 服务 {#the-service}

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

`client` 和 `state` 是 Angular 信号；`disconnect()`（销毁时也会调用）会释放客户端。
