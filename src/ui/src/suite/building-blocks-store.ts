import type { WordBuildingBlocksStore } from '../docx/building-blocks-shared';

/**
 * Word's shared building blocks (Building Blocks.dotx) for one profile, in an IndexedDB database
 * of that name on this device. Without IndexedDB the entries last until the page closes. The
 * IndexedDB code (`building-blocks-idb.ts`) downloads on the first read or write.
 */
export function createBuildingBlocksStore(name: string): WordBuildingBlocksStore {
	let store: Promise<WordBuildingBlocksStore> | undefined;
	const load = () =>
		(store ??= import('./building-blocks-idb').then((idb) => idb.createBuildingBlocksStore(name)));
	return {
		read: async () => (await load()).read(),
		write: async (value) => (await load()).write(value),
	};
}
