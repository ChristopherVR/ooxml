// Connects a Lit element to a core `TeamsClient`: creates it, re-renders the host whenever the
// state changes and tears it down. This is the whole "framework glue" of the element; the bindings
// for React, Vue, Solid, Svelte and Angular do the same few lines with their own reactivity.
import type { ReactiveController, ReactiveControllerHost } from 'lit';
import {
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsState,
	createTeamsClient,
} from 'ooxml-core/teams';

export class TeamsController implements ReactiveController {
	client: TeamsClient | null = null;
	state: TeamsState | null = null;
	private off: (() => void)[] = [];

	constructor(
		private readonly host: ReactiveControllerHost,
		private readonly onNotice: (text: string) => void,
	) {
		host.addController(this);
	}

	start(options: TeamsClientOptions): void {
		this.stop();
		const client = createTeamsClient(options);
		this.client = client;
		this.state = client.getState();
		this.off = [
			client.subscribe(() => {
				this.state = client.getState();
				this.host.requestUpdate();
			}),
			client.on('notice', this.onNotice),
		];
		this.host.requestUpdate();
	}

	stop(): void {
		for (const fn of this.off.splice(0)) fn();
		this.client?.destroy();
		this.client = null;
		this.state = null;
	}

	hostDisconnected(): void {
		this.stop();
	}
}
