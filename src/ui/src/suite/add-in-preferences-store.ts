import type { OfficeAddInPreferences } from '../ribbon/add-in-preferences';

/**
 * A local manifest store shared by the suite and the demos. Its IndexedDB code
 * (`add-in-preferences-idb.ts`) downloads on the first read or write.
 */
export function createAddInPreferencesStore(name: string): OfficeAddInPreferences {
	let store: Promise<OfficeAddInPreferences> | undefined;
	const load = () =>
		(store ??= import('./add-in-preferences-idb').then((idb) =>
			idb.createAddInPreferencesStore(name),
		));
	return {
		read: async (host) => (await load()).read(host),
		write: async (host, id, bytes) => (await load()).write(host, id, bytes),
	};
}
