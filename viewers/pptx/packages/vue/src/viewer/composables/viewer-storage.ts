import type { DocumentStorage } from 'ooxml-ui/pptx';
/**
 * The host's `storage` prop for deep chrome (File > Open > Recent, Account >
 * Storage & Privacy), so they list and read recovery copies from the host's
 * store when it has one. `PowerPointViewer.vue` provides a getter; a component
 * rendered standalone sees none (the library's IndexedDB).
 */
import { inject } from 'vue';
import type { InjectionKey } from 'vue';

export const ViewerStorageKey: InjectionKey<() => DocumentStorage | undefined> =
	Symbol('pptx-viewer-storage');

/** A getter for the viewer's host-owned storage (undefined when the host passed none). */
export function useViewerStorage(): () => DocumentStorage | undefined {
	return inject(ViewerStorageKey, () => undefined);
}
