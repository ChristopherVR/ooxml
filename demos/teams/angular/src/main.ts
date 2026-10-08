import '@angular/compiler';
import { Component, signal } from '@angular/core';
import { bootstrapApplication } from '@angular/platform-browser';
import { TeamsWorkspaceComponent } from 'openteams-angular-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';
import { currentHostClass, onHostClass, recordOpenFile } from '../../test-hooks';

// JIT-compiled in the demo: legacy decorators, no AOT.
@Component({
	selector: 'demo-root',
	standalone: true,
	imports: [TeamsWorkspaceComponent],
	template: `<teams-workspace
		[className]="hostClass()"
		[workspaceId]="workspaceId"
		[userName]="userName"
		[userId]="userId"
		[config]="config"
		(openFile)="openFile($event)"
	/>`,
	styles: [':host { display: block; height: 100%; }'],
})
class DemoRoot {
	workspaceId = workspaceId;
	userName = userName;
	userId = userId;
	config = config;
	// The host class of <teams-app>; the browser tests swap it (../../test-hooks).
	hostClass = signal(currentHostClass());
	constructor() {
		onHostClass((value) => this.hostClass.set(value));
	}
	// Angular outputs carry one value: `{ detail, event }`.
	openFile({ detail, event }: { detail: Parameters<typeof recordOpenFile>[0]; event: Event }) {
		recordOpenFile(detail, event);
	}
}

showStaticNotice();
void bootstrapApplication(DemoRoot);
