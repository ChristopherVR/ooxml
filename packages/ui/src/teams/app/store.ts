// The raw, framework-neutral hook surface. Every binding wraps exactly this: create a client once,
// then read `getState()` and call its actions. Use it to build your own UI without <teams-app>.
//
//   const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
//   const stop = teams.subscribe(() => render(teams.getState()));
//   teams.send({ text: 'hello' });
import { type TeamsClient, type TeamsClientOptions, createTeamsClient } from 'ooxml-core/teams';
import { safeStorage } from './storage.js';

export {
	type CallView,
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsState,
	type ChannelView,
	type PersonView,
	type FileEntry,
	type SearchHit,
} from 'ooxml-core/teams';

/** `createTeamsClient` with the browser's safe localStorage wired in for offline snapshots. */
export function createTeams(options: TeamsClientOptions): TeamsClient {
	return createTeamsClient({ storage: safeStorage, ...options });
}
