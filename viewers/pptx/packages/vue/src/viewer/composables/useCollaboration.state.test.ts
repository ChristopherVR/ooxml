import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { afterEach, describe, expect, it } from 'vitest';
import { effectScope, nextTick, ref, shallowRef } from 'vue';

import { useCollaboration } from './useCollaboration';

const cleanup: (() => void)[] = [];
afterEach(() => {
	for (const dispose of cleanup.splice(0).reverse()) dispose();
});

describe('useCollaboration state (the shared CollaborationState)', () => {
	it('reports a host-owned session with its people, then the off state on stop', async () => {
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
		>({ session, user: { name: 'Bo' } });
		const scope = effectScope();
		cleanup.push(() => scope.stop());
		const collab = scope.run(() =>
			useCollaboration({
				collaboration,
				slides: ref([]),
				onRemoteSlides: () => {},
				loadVersion: ref(0),
				getLoadOrigin: () => 'bootstrap',
			} as never),
		)!;
		await nextTick();
		await new Promise((resolve) => setTimeout(resolve, 0));
		session.awareness.states.set(99, {
			presence: {
				clientId: 99,
				userName: 'Cy',
				userColor: '#ef4444',
				activeSlideIndex: 0,
				cursorX: 1,
				cursorY: 1,
				lastUpdated: new Date().toISOString(),
			},
		});
		session.awareness.emit('change', [{ added: [99], updated: [], removed: [] }, 'remote']);
		await nextTick();
		const state = collab.state.value;
		expect(state).toMatchObject({ active: true, roomId: 'shared-room' });
		expect(state.status).not.toBe('off');
		expect(state.people).toEqual([
			expect.objectContaining({ clientId: session.awareness.clientID, name: 'Bo', self: true }),
			expect.objectContaining({ clientId: 99, name: 'Cy', self: false }),
		]);
		collaboration.value = undefined;
		await nextTick();
		expect(collab.state.value).toMatchObject({
			active: false,
			roomId: undefined,
			status: 'off',
			synced: false,
			people: [],
		});
	});

	it('reports a write-back status for an owner with onWriteBack only', async () => {
		const hub = createMemoryHub();
		const session = createCollabSession({
			roomId: 'wb-room',
			provider: transportProvider({ transport: hub.createTransport('wb-room') }),
			user: { name: 'Ada' },
			heartbeatMs: 0,
		});
		cleanup.push(() => session.destroy());
		const collaboration = shallowRef<Record<string, unknown> | undefined>({
			session,
			user: { name: 'Ada', role: 'owner' },
			onWriteBack: () => {},
		});
		const scope = effectScope();
		cleanup.push(() => scope.stop());
		const collab = scope.run(() =>
			useCollaboration({
				collaboration,
				slides: ref([]),
				onRemoteSlides: () => {},
				loadVersion: ref(0),
				getLoadOrigin: () => 'bootstrap',
			} as never),
		)!;
		await nextTick();
		await new Promise((resolve) => setTimeout(resolve, 0));
		expect(collab.state.value.writeBack).toEqual({ status: 'idle' });
		collaboration.value = {
			session,
			user: { name: 'Ada', role: 'collaborator' },
			onWriteBack: () => {},
		};
		await nextTick();
		await new Promise((resolve) => setTimeout(resolve, 0));
		expect(collab.state.value.writeBack).toBeUndefined();
	});
});
