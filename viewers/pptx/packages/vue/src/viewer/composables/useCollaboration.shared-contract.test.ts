import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { afterEach, describe, expect, it } from 'vitest';
import { effectScope, nextTick, ref, shallowRef } from 'vue';

import { useCollaboration } from './useCollaboration';

const cleanup: (() => void)[] = [];
afterEach(() => {
	for (const dispose of cleanup.splice(0).reverse()) dispose();
});

describe('useCollaboration with the shared collaboration contract', () => {
	it('borrows a host CollabSession passed as collaboration.session and never destroys it', async () => {
		const hub = createMemoryHub();
		const session = createCollabSession({
			roomId: 'shared-room',
			provider: transportProvider({ transport: hub.createTransport('shared-room') }),
			user: { name: 'Ada' },
			heartbeatMs: 0,
		});
		cleanup.push(() => session.destroy());
		const collaboration = shallowRef<
			{ session: typeof session; user: { name: string } } | undefined
		>({
			session,
			user: { name: 'Bo' },
		});
		const scope = effectScope();
		cleanup.push(() => scope.stop());
		const slides = ref([]);
		const collab = scope.run(() =>
			useCollaboration({
				collaboration,
				slides,
				onRemoteSlides: () => {},
				loadVersion: ref(0),
				getLoadOrigin: () => 'bootstrap',
			} as never),
		)!;
		await nextTick();
		await new Promise((resolve) => setTimeout(resolve, 0));
		expect(collab.active.value).toBeTruthy();
		expect(collab.activeCollaboration.value?.roomId).toBe('shared-room');
		expect(collab.activeCollaboration.value?.userName).toBe('Bo');
		expect(collab.activeCollaboration.value?.externalSession?.doc).toBe(session.doc);
		collaboration.value = undefined;
		await nextTick();
		expect(collab.active.value).toBeFalsy();
		expect(session.doc.isDestroyed).toBeFalsy();
	});
});
