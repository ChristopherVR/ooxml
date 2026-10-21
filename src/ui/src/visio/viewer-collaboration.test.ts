import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	COLLABORATION_CHANGE_EVENT,
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollaborationState,
} from 'ooxml-core/collab';
import { mountViewer } from './binding';

afterEach(() => document.body.replaceChildren());

const session = (name: string) =>
	createCollabSession({
		roomId: 'element',
		provider: transportProvider({ transport: createMemoryHub().createTransport('element') }),
		user: { name },
		heartbeatMs: 0,
		syncGraceMs: 0,
	});

describe('<visio-viewer> collaboration', () => {
	it('shares while `collaboration` is set, reports state and leaves on null', async () => {
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host, {});
		const { element } = viewer;
		const heard: CollaborationState[] = [];
		element.addEventListener(COLLABORATION_CHANGE_EVENT, (event) =>
			heard.push((event as CustomEvent<CollaborationState>).detail),
		);
		const mine = session('Ada');
		const destroy = vi.spyOn(mine, 'destroy');
		expect(element.collaborationState).toMatchObject({ active: false, status: 'off', people: [] });
		const options = { session: mine };
		element.collaboration = options;
		expect(element.collaboration).toBe(options);
		await vi.waitFor(() => expect(element.collaborationState.active).toBe(true));
		expect(heard.at(-1)).toMatchObject({ active: true, roomId: 'element' });
		expect(heard.at(-1)?.people[0]).toMatchObject({ name: 'Ada', self: true });

		element.reconnectCollaboration();
		element.collaboration = null;
		expect(element.collaborationState.active).toBe(false);
		expect(heard.at(-1)).toMatchObject({ active: false, status: 'off' });
		expect(destroy).not.toHaveBeenCalled();

		element.collaboration = options;
		await vi.waitFor(() => expect(element.collaborationState.active).toBe(true));
		element.stopCollaboration();
		expect(element.collaboration).toBeNull();
		expect(element.collaborationState.active).toBe(false);
		viewer.destroy();
		expect(destroy).not.toHaveBeenCalled();
		mine.destroy();
	});

	it('destroys a session the viewer opened itself and rejoins after a reconnect of the element', async () => {
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host, { collaboration: { roomId: 'own-room' } });
		const { element } = viewer;
		await vi.waitFor(() => expect(element.collaborationState.roomId).toBe('own-room'));
		element.remove();
		expect(element.collaborationState.active).toBe(false);
		expect(element.collaboration).toEqual({ roomId: 'own-room' });
		host.append(element);
		await vi.waitFor(() => expect(element.collaborationState.active).toBe(true));
		viewer.destroy();
		expect(element.collaborationState.active).toBe(false);
	});
});
