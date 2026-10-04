import { Injectable, type OnDestroy, signal } from '@angular/core';
import {
	createTeams,
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsState,
} from 'teams-viewer';

/**
 * The raw service: provide it in a component (`providers: [TeamsService]`), call `connect()`
 * once and read `state()` (an Angular signal) in templates. Actions are on `client()`.
 *
 *   svc.connect({ workspaceId: 'acme', user, config });
 *   svc.state()?.channels;   svc.client()?.send({ text: 'hi' });
 */
@Injectable()
export class TeamsService implements OnDestroy {
	readonly client = signal<TeamsClient | null>(null);
	readonly state = signal<TeamsState | null>(null);
	private off: (() => void) | null = null;

	connect(options: TeamsClientOptions): TeamsClient {
		this.disconnect();
		const created = createTeams(options);
		this.client.set(created);
		this.state.set(created.getState());
		this.off = created.subscribe(() => this.state.set(created.getState()));
		return created;
	}

	disconnect(): void {
		this.off?.();
		this.off = null;
		this.client()?.destroy();
		this.client.set(null);
		this.state.set(null);
	}

	ngOnDestroy(): void {
		this.disconnect();
	}
}
