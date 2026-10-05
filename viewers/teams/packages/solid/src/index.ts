import { createEffect, createSignal, onCleanup, onMount, type Accessor } from 'solid-js';
import {
	applyTeamsProps,
	createTeams,
	defineTeamsApp,
	listenTeamsEvents,
	type TeamsApp,
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsProps,
	type TeamsState,
} from 'teams-viewer';

export type { TeamsClient, TeamsClientOptions, TeamsProps, TeamsState } from 'teams-viewer';

/**
 * Returns the `<teams-app>` element; use it as `{Teams(props)}` inside JSX. A plain function (no
 * JSX transform needed) keeps the package a lifecycle adapter: props are re-applied reactively
 * and handlers are read at event time.
 */
export function Teams(props: TeamsProps): HTMLElement {
	defineTeamsApp();
	const el = document.createElement('teams-app') as TeamsApp;
	el.style.display = 'block';
	el.style.height = '100%';
	let applied: TeamsProps = {};
	const snapshot = (): TeamsProps => ({
		...(props.workspaceId !== undefined ? { workspaceId: props.workspaceId } : {}),
		...(props.userName !== undefined ? { userName: props.userName } : {}),
		...(props.userId !== undefined ? { userId: props.userId } : {}),
		...(props.config !== undefined ? { config: props.config } : {}),
		...(props.uploadFile ? { uploadFile: props.uploadFile } : {}),
		...(props.openers ? { openers: props.openers } : {}),
	});
	createEffect(() => {
		const next = snapshot();
		applyTeamsProps(el, next, applied);
		applied = next;
	});
	onMount(() => onCleanup(listenTeamsEvents(el, () => props)));
	return el;
}

/**
 * The raw primitive: a client for `options()`, recreated when it changes and destroyed with the
 * owner. `state()` is a signal that holds the latest snapshot.
 *
 *   const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user, config }));
 */
export function createTeamsClient(options: () => TeamsClientOptions | null): {
	client: Accessor<TeamsClient | null>;
	state: Accessor<TeamsState | null>;
} {
	const [client, setClient] = createSignal<TeamsClient | null>(null);
	const [state, setState] = createSignal<TeamsState | null>(null, { equals: false });
	createEffect(() => {
		const opts = options();
		if (!opts) return;
		const created = createTeams(opts);
		setClient(created);
		setState(created.getState());
		const off = created.subscribe(() => setState(created.getState()));
		onCleanup(() => {
			off();
			created.destroy();
			setClient(null);
			setState(null);
		});
	});
	return { client, state };
}
