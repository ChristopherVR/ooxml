import { describe, expect, it } from 'vitest';
import { ref, shallowRef } from 'vue';

import { useSaveRevision } from './useSaveRevision';
import { useViewerApi } from './useViewerApi';
import type { UseViewerApiOptions } from './useViewerApi';

function setup() {
	const slides = shallowRef<unknown[]>([]);
	const isDirty = ref(true);
	const saveRevision = useSaveRevision({
		changeSources: () => [slides.value],
		getContent: async () => new Uint8Array([7]),
		markClean: () => {
			isDirty.value = false;
		},
		markDirty: () => {
			isDirty.value = true;
		},
		filePath: () => undefined,
		storage: () => undefined,
		getOptions: () => undefined,
	});
	const api = useViewerApi({ saveRevision, isDirty } as unknown as UseViewerApiOptions);
	return { api, slides, isDirty };
}

describe('Vue host-persisted save acknowledgement', () => {
	it('cleans the deck only for the current revision', async () => {
		const { api, slides, isDirty } = setup();
		const snapshot = await api.getSaveSnapshot();
		expect([...snapshot.content]).toEqual([7]);
		expect(isDirty.value).toBe(true);
		slides.value = [{ id: 'edited during save' }];
		expect(api.acknowledgeSave(snapshot.revision).status).toBe('pending');
		expect(isDirty.value).toBe(true);
		const latest = await api.getSaveSnapshot();
		expect(api.getSaveRevision()).toBe(latest.revision);
		expect(api.acknowledgeSave(latest.revision).status).toBe('clean');
		expect(isDirty.value).toBe(false);
	});

	it('ignores a stale or unknown revision', async () => {
		const { api, slides, isDirty } = setup();
		const first = await api.getSaveSnapshot();
		slides.value = [{}];
		const second = await api.getSaveSnapshot();
		expect(api.acknowledgeSave(second.revision).status).toBe('clean');
		isDirty.value = true;
		slides.value = [{}, {}];
		expect(api.acknowledgeSave(first.revision).status).toBe('stale');
		expect(api.acknowledgeSave(12345).status).toBe('unknown');
		expect(isDirty.value).toBe(true);
	});

	it('stays pending and dirty when a recovery snapshot cleared the flag after a newer edit', async () => {
		const { api, slides, isDirty } = setup();
		const snapshot = await api.getSaveSnapshot();
		slides.value = [{ id: 'edited after the host snapshot' }];
		// The autosave engine writes a recovery snapshot and clears its (shared) dirty flag.
		isDirty.value = false;
		expect(api.acknowledgeSave(snapshot.revision).status).toBe('pending');
		expect(isDirty.value).toBe(true);
	});
});
