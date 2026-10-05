import { readable, type Readable } from 'svelte/store';
import {
	createTeams,
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsState,
} from 'teams-viewer';

export type { TeamsClient, TeamsClientOptions, TeamsState } from 'teams-viewer';

/**
 * The raw store: subscribe to get the latest state snapshot (`null` until connected). The client
 * lives as long as the store has subscribers; its actions are on `client()`.
 *
 *   const teams = teamsStore({ workspaceId: 'acme', user, config });
 *   $teams?.channels   // reactive
 *   teams.client()?.send({ text: 'hi' })
 */
export function teamsStore(options: TeamsClientOptions): Readable<TeamsState | null> & {
	client: () => TeamsClient | null;
} {
	let current: TeamsClient | null = null;
	const store = readable<TeamsState | null>(null, (set) => {
		const created = createTeams(options);
		current = created;
		set(created.getState());
		const off = created.subscribe(() => set(created.getState()));
		return () => {
			off();
			created.destroy();
			current = null;
		};
	});
	return { subscribe: store.subscribe, client: () => current };
}
