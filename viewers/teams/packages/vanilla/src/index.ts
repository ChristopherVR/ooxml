import {
	applyTeamsProps,
	defineTeamsApp,
	listenTeamsEvents,
	type TeamsApp,
	type TeamsProps,
} from 'teams-viewer';

// The raw store and the element itself, for your own UI or plain `<teams-app>` markup: the private
// `teams-viewer` package is inlined into this one when it is built, so re-export what users need.
export {
	TeamsApp,
	createTeams,
	defineTeamsApp,
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsProps,
	type TeamsState,
} from 'teams-viewer';

export interface MountedTeams {
	readonly element: TeamsApp;
	/** Update props; only changed values are applied. */
	update: (next: TeamsProps) => void;
	destroy: () => void;
}

/** Create a `<teams-app>` inside `container` and keep it in sync with `props`. */
export function mountTeams(container: HTMLElement, props: TeamsProps = {}): MountedTeams {
	defineTeamsApp();
	const element = container.ownerDocument.createElement('teams-app') as TeamsApp;
	element.style.display = 'block';
	element.style.height = '100%';
	let current = props;
	const stop = listenTeamsEvents(element, () => current);
	applyTeamsProps(element, current);
	container.append(element);
	return {
		element,
		update(next) {
			applyTeamsProps(element, next, current);
			current = next;
		},
		destroy() {
			stop();
			element.remove();
		},
	};
}
