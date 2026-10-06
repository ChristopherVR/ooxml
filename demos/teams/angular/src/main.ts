import '@angular/compiler';
import { Component } from '@angular/core';
import { bootstrapApplication } from '@angular/platform-browser';
import { TeamsWorkspaceComponent } from 'openteams-angular-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';

// JIT-compiled in the demo: legacy decorators, no AOT.
@Component({
	selector: 'demo-root',
	standalone: true,
	imports: [TeamsWorkspaceComponent],
	template: `<teams-workspace
		[workspaceId]="workspaceId"
		[userName]="userName"
		[userId]="userId"
		[config]="config"
	/>`,
	styles: [':host { display: block; height: 100%; }'],
})
class DemoRoot {
	workspaceId = workspaceId;
	userName = userName;
	userId = userId;
	config = config;
}

showStaticNotice();
void bootstrapApplication(DemoRoot);
