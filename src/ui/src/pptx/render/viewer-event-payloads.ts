/**
 * @fileoverview The payloads of the viewer events every binding reports the same way.
 *
 * Each binding keeps its own autosave engine and so its own internal status shape (a bare string
 * in Vue and Svelte, a `{ state }` object in React and Angular, a string without `disabled` in
 * Vanilla). The host-facing `autosaveStatus` event must not leak that difference, so every
 * binding maps what it holds through {@link autosaveStatusKind} before reporting it.
 *
 * @module render/viewer-event-payloads
 */

/** The autosave lifecycle a host is told about (`autosaveStatus`). */
export type ViewerAutosaveStatus = 'idle' | 'saving' | 'saved' | 'error';

/**
 * Collapse a binding's internal autosave status to the public lifecycle. A disabled engine (the
 * host forbade autosave, the user switched it off, or there is no file key) reports `idle`: no
 * save is in flight and none will start.
 */
export function autosaveStatusKind(
	status: string | { readonly state: string },
): ViewerAutosaveStatus {
	const state = typeof status === 'string' ? status : status.state;
	return state === 'saving' || state === 'saved' || state === 'error' ? state : 'idle';
}
