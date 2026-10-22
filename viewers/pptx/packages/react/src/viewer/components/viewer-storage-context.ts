import type { DocumentStorage } from 'ooxml-ui/pptx';
import { createContext, useContext } from 'react';

/**
 * Shares the host's `storage` prop with deep chrome components (File > Open >
 * Recent, Account > Storage & Privacy), so they list and read recovery copies
 * from the host's store when it has one. `PowerPointViewer` provides it;
 * standalone consumers see none (the library's IndexedDB).
 */
export const ViewerStorageContext = createContext<DocumentStorage | undefined>(undefined);

/** The viewer's host-owned storage, if the host passed one. */
export function useViewerStorage(): DocumentStorage | undefined {
	return useContext(ViewerStorageContext);
}
