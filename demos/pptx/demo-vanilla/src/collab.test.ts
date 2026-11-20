/**
 * A demo opened from a peer-to-peer link that lacks its `#key=` secret (an older
 * link kept the key in the query) must not silently stay offline: the page
 * address still names the room, and the viewer's session fails with the
 * missing-key reason its status pill and Share dialog translate.
 */
import {
	createCollaborationController,
	createInitialViewerState,
	createStore,
} from 'pptx-vanilla-viewer';
import { afterEach, describe, expect, it } from 'vitest';

import { sessionFromPage } from '../../shared/collaboration-link';
import { buildRoomConfig, readRoomFromUrl } from './collab';

const SECRET = 'AAECAwQFBgcICQoLDA0ODw';
const original = window.location.href;

function openPage(address: string): void {
	window.history.replaceState({}, '', address);
}

afterEach(() => openPage(original));

describe('demo session from the page address', () => {
	it('names a keyless peer-to-peer room without a secret', () => {
		for (const address of [
			'/?room=deck-1&transport=webrtc',
			`/?room=deck-1&transport=webrtc&key=${SECRET}`,
		]) {
			openPage(address);
			expect(sessionFromPage('room', '')).toEqual({
				roomId: 'deck-1',
				serverUrl: '',
				transport: 'webrtc',
			});
		}
	});

	it('keeps the secret of a complete link and ignores a page without a room', () => {
		openPage(`/?room=deck-1&transport=webrtc#key=${SECRET}`);
		expect(sessionFromPage('room', '')).toEqual({
			roomId: 'deck-1',
			serverUrl: '',
			transport: 'webrtc',
			secret: SECRET,
		});
		openPage('/?sample=1');
		expect(sessionFromPage('room', '')).toBeNull();
	});

	it('starts a keyless link into the missing-key failure', async () => {
		openPage('/?room=deck-1&transport=webrtc');
		const link = readRoomFromUrl();
		expect(link).not.toBeNull();
		const collab = createCollaborationController({
			store: createStore(createInitialViewerState()),
			getHandler: () => null,
			setEditable: () => {},
		});
		await collab.start(buildRoomConfig(link!, 'Ada'));
		expect(collab.getStatus()).toBe('error');
		expect(collab.getFailure()).toBe('missing-secret');
		collab.destroy();
	});
});
