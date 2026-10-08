// Hooks the browser tests (e2e/teams) drive in every framework demo, so one spec checks each
// binding the same way: the host class every demo passes to its binding, a way to swap it, and a
// record of the binding's `open-file` callback. They do nothing until a test calls them.

/** The host class every demo gives the workspace element through its binding. */
export const HOST_CLASS = 'demo-teams-host';

/** What a binding's `open-file` handler received, kept for the tests to read. */
export interface OpenedFile {
	name: string;
	kind: string;
	url: string | undefined;
	/** The type of the event the binding passed with the detail. */
	eventType: string | undefined;
}

interface DemoHooks {
	setHostClass(value: string): void;
	openedFiles: OpenedFile[];
}

declare global {
	interface Window {
		teamsDemo?: DemoHooks;
	}
}

let hostClass = HOST_CLASS;
const listeners = new Set<(value: string) => void>();
const hooks: DemoHooks = {
	setHostClass(value) {
		hostClass = value;
		for (const listener of listeners) listener(value);
	},
	openedFiles: [],
};
window.teamsDemo = hooks;

/** The current host class. */
export const currentHostClass = (): string => hostClass;

/** Call `listener` whenever a test swaps the host class; returns the unsubscribe function. */
export function onHostClass(listener: (value: string) => void): () => void {
	listeners.add(listener);
	return () => {
		listeners.delete(listener);
	};
}

/** The demos' `open-file` handler: record what the binding passed. */
export function recordOpenFile(
	detail: { attachment: { name: string; kind: string }; url: string | undefined },
	event?: Event,
): void {
	hooks.openedFiles.push({
		name: detail.attachment.name,
		kind: detail.attachment.kind,
		url: detail.url,
		eventType: event?.type,
	});
}
