import type { DocumentStorage, ViewerOptions } from 'ooxml-ui/pptx';
import {
	clearAutosaveSnapshots,
	pruneExpiredAutosaveSnapshots,
	shouldClearAutosaveCacheOnClose,
} from 'ooxml-ui/pptx';
import { useCallback, useEffect, useRef } from 'react';

// Keep this callback outside the viewer's render scope: browser listeners must
// not retain a render's full slide state through a shared lexical context.
function clearOptionsCache(storage: DocumentStorage | undefined): void {
	void clearAutosaveSnapshots(storage).catch(() => {
		// Browser lifecycle cleanup cannot report a storage failure.
	});
}

/**
 * Full-viewer cache maintenance, separate from the headless options store. With a host
 * recovery store (`storage`), the sweeps run against it only when it lists its copies;
 * otherwise retention is the host's job (see the shared `autosave-maintenance`).
 */
export function useAutosaveCacheMaintenance(
	viewerOptions: ViewerOptions,
	storage?: DocumentStorage,
): () => void {
	const storageRef = useRef(storage);
	storageRef.current = storage;
	// File > Options > Save > "cache retention": a one-time sweep per mount is
	// enough, since a fresh snapshot only ever lands with a fresh timestamp.
	useEffect(() => {
		void pruneExpiredAutosaveSnapshots(viewerOptions, storageRef.current).catch(() => {
			// Best-effort background maintenance; a blocked store skips it.
		});
		// eslint-disable-next-line react-hooks/exhaustive-deps -- one sweep per mount, not per option edit
	}, []);

	// File > Options > Save > "clear cache on close": wipe recovery snapshots
	// when the tab closes/navigates away, and when this viewer unmounts.
	useEffect(() => {
		const clearIfRequested = (): void => {
			if (shouldClearAutosaveCacheOnClose(viewerOptions)) {
				clearOptionsCache(storageRef.current);
			}
		};
		window.addEventListener('beforeunload', clearIfRequested);
		return () => {
			window.removeEventListener('beforeunload', clearIfRequested);
			clearIfRequested();
		};
	}, [viewerOptions]);

	const clearHostCache = useCallback(() => clearOptionsCache(storageRef.current), []);
	// Without a host store the callback is one module-level function, stable across mounts.
	return storage ? clearHostCache : clearBrowserCache;
}

const clearBrowserCache = (): void => clearOptionsCache(undefined);
