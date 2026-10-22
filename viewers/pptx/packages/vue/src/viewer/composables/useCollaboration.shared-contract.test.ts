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

	function mount(collaboration: ReturnType<typeof shallowRef>) {
		const scope = effectScope();
		cleanup.push(() => scope.stop());
		return scope.run(() =>
			useCollaboration({
				collaboration,
				slides: ref([]),
				onRemoteSlides: () => {},
				loadVersion: ref(0),
				getLoadOrigin: () => 'bootstrap',
			} as never),
		)!;
	}
	const settle = async () => {
		await nextTick();
		await new Promise((resolve) => setTimeout(resolve, 20));
	};

	it('opens a session around a provider-only config and destroys it on stop', async () => {
		const hub = createMemoryHub();
		const host = createCollabSession({
			roomId: 'owned-room',
			provider: transportProvider({ transport: hub.createTransport('owned-room') }),
			user: { name: 'Host' },
			heartbeatMs: 0,
		});
		cleanup.push(() => host.destroy());
		const collaboration = shallowRef<object | undefined>({
			roomId: 'owned-room',
			provider: transportProvider({ transport: hub.createTransport('owned-room') }),
			user: { name: 'Ada' },
			intent: 'join',
		});
		const collab = mount(collaboration);
		expect(collab.status.value).toBe('connecting');
		await settle();
		expect(collab.active.value).toBeTruthy();
		expect(collab.activeCollaboration.value?.roomId).toBe('owned-room');
		expect(collab.state.value).toMatchObject({ active: true, roomId: 'owned-room' });
		expect(host.peers()).toHaveLength(1);
		collaboration.value = undefined;
		await nextTick();
		expect(collab.active.value).toBeFalsy();
		await settle();
		expect(host.peers()).toHaveLength(0);
		expect(host.doc.isDestroyed).toBeFalsy();
	});

	it('shares between windows without a server and opens nothing when stopped while pending', async () => {
		const collaboration = shallowRef<object | undefined>({
			roomId: 'tabs',
			user: { name: 'Tab' },
		});
		const collab = mount(collaboration);
		await settle();
		expect(collab.status.value).toBe('connected');
		collaboration.value = { roomId: 'pending', user: { name: 'Tab' } };
		await nextTick();
		collaboration.value = undefined;
		await settle();
		expect(collab.active.value).toBeFalsy();
		expect(collab.status.value).toBe('disconnected');
	});
});
