/**
 * A demo opened from a peer-to-peer link that lacks its `#key=` secret (an older
 * link kept the key in the query) must not silently stay offline: the page
 * address still names the room, and the viewer's session fails with the
 * missing-key reason its status pill and Share dialog translate. A `?broadcast=`
 * link joins the audience, and a presenter's address bar shows that link.
 */
import {
	createCollaborationController,
	createInitialViewerState,
	createStore,
} from 'pptx-vanilla-viewer';
import { afterEach, describe, expect, it } from 'vitest';

import { sessionFromPage, showSessionInAddressBar } from '../../shared/collaboration-link';
import { buildInvitationConfig, readInvitationFromUrl } from './collab';

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
		const invitation = readInvitationFromUrl();
		expect(invitation).not.toBeNull();
		const collab = createCollaborationController({
			store: createStore(createInitialViewerState()),
			getHandler: () => null,
			setEditable: () => {},
		});
		await collab.start(buildInvitationConfig(invitation!, 'Ada'));
		expect(collab.getStatus()).toBe('error');
		expect(collab.getFailure()).toBe('missing-secret');
		collab.destroy();
	});

	it('joins a broadcast link as the audience, which follows the presenter', () => {
		openPage(`/?broadcast=show-1&transport=webrtc#key=${SECRET}`);
		const invitation = readInvitationFromUrl();
		expect(invitation).not.toBeNull();
		expect(buildInvitationConfig(invitation!, 'Ada')).toMatchObject({
			roomId: 'show-1',
			transport: 'webrtc',
			secret: SECRET,
			role: 'viewer',
			sessionIntent: 'join',
		});
	});

	it('joins a room link without seeding it, unless the page brings its own deck', () => {
		openPage(`/?room=deck-1&transport=webrtc#key=${SECRET}`);
		const joining = buildInvitationConfig(readInvitationFromUrl()!, 'Ada');
		expect(joining).toMatchObject({ roomId: 'deck-1', sessionIntent: 'join' });
		expect(joining.role).toBeUndefined();
		openPage(`/?sample=1&room=deck-1&transport=webrtc#key=${SECRET}`);
		expect(buildInvitationConfig(readInvitationFromUrl()!, 'Ada').sessionIntent).toBe('create');
	});

	it('shows a broadcast under its audience parameter in the address bar', () => {
		openPage('/?sample=1');
		showSessionInAddressBar({
			roomId: 'show-1',
			serverUrl: '',
			transport: 'webrtc',
			secret: SECRET,
			role: 'owner',
		});
		const address = new URL(window.location.href);
		expect(address.searchParams.get('broadcast')).toBe('show-1');
		expect(address.searchParams.get('room')).toBeNull();
		expect(address.searchParams.get('sample')).toBe('1');
	});
});
