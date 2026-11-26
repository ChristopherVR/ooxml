import type { AddInDialogHost } from './add-in-dialog';

/** A selection bound to the document that was current when the pane opened. */
export interface OfficeAddInSelection {
	readText(): string | undefined;
	writeText(text: string): boolean;
}

export interface OfficeAddInRuntime {
	/** Approved publisher-controlled HTTPS origins. */
	origins: readonly string[];
	/**
	 * The embedding page enforces a Content Security Policy whose frame-src restricts
	 * frame navigations, including redirects, to the approved publisher origins.
	 * Set this only after applying that policy. The runtime blocks panes without it.
	 */
	frameSrcEnforced: true;
	/** Optional adapter for a host with its own live document integration. */
	selection?: () => OfficeAddInSelection | undefined;
}

const RUNTIME = Symbol.for('office-ui.add-in-runtime');
export const ADD_IN_RUNTIME_CHANGED = 'office-add-in-runtime-changed';
export const ADD_IN_RUNTIME_REFRESHED = 'office-add-in-runtime-refreshed';
type Settings = AddInDialogHost & { [RUNTIME]?: OfficeAddInRuntime | null };

/** Opt into compatible task panes in this scope. Passing null disables inherited access. */
export function setOfficeAddInRuntime(host: AddInDialogHost, runtime: OfficeAddInRuntime | null) {
	const target: Settings = host;
	target[RUNTIME] = runtime;
	target.dispatchEvent(new Event(ADD_IN_RUNTIME_CHANGED));
}

export function addInRuntime(scopes: readonly AddInDialogHost[]) {
	for (const scope of scopes) {
		const settings: Settings = scope;
		if (settings[RUNTIME] !== undefined) return settings[RUNTIME];
	}
}

export function allowedPaneUrl(source: string, runtime: OfficeAddInRuntime | null | undefined) {
	try {
		const url = new URL(source);
		if (
			url.protocol === 'https:' &&
			!url.username &&
			!url.password &&
			runtime?.frameSrcEnforced === true &&
			runtime.origins.includes(url.origin)
		)
			return url;
	} catch {
		// Configuration and manifests may contain unusable URLs.
	}
}
