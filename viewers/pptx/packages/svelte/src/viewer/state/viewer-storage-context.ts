import type { DocumentStorage } from 'ooxml-ui/pptx';
import { getContext, setContext } from 'svelte';

/**
 * Svelte context for the host's `storage` prop, so deep chrome (File > Open >
 * Recent, Account > Storage & Privacy) lists and reads recovery copies from the
 * host's store when it has one. `createViewerState` provides a getter; a
 * component mounted stand-alone sees none (the library's IndexedDB).
 */
const VIEWER_STORAGE_CONTEXT_KEY = Symbol('pptx-svelte-viewer-storage');

/** Provide the storage getter to the component subtree (root component only). */
export function provideViewerStorage(getStorage: () => DocumentStorage | undefined): void {
	setContext(VIEWER_STORAGE_CONTEXT_KEY, getStorage);
}

/** A getter for the viewer's host-owned storage (undefined when the host passed none). */
export function useViewerStorage(): () => DocumentStorage | undefined {
	return (
		getContext<(() => DocumentStorage | undefined) | undefined>(VIEWER_STORAGE_CONTEXT_KEY) ??
		(() => undefined)
	);
}
