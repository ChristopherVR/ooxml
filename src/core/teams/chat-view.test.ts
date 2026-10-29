import { contactViews, memberPresence, peerPresenceOf } from './chat-view';
import type { PeerLike } from './view';

const peer = (clientId: number, userId: string, status: string, statusMessage = ''): PeerLike => ({
	clientId,
	userName: userId,
	userColor: '#000',
	userId,
	status,
	...(statusMessage ? { statusMessage } : {}),
});

describe('presence of a person with two tabs', () => {
	const idleFirst = [peer(1, 'ada', 'away'), peer(2, 'ada', 'available', 'at my desk')];
	const activeFirst = [...idleFirst].reverse();

	it('shows the more present tab whichever connected first', () => {
		expect(memberPresence('ada', 'me', 'available', idleFirst)).toBe('available');
		expect(memberPresence('ada', 'me', 'available', activeFirst)).toBe('available');
	});

	it('shows busy from any tab over an available one', () => {
		const peers = [peer(1, 'ada', 'available'), peer(2, 'ada', 'busy')];
		expect(memberPresence('ada', 'me', 'available', peers)).toBe('busy');
		expect(memberPresence('ada', 'me', 'available', [...peers].reverse())).toBe('busy');
	});

	it('takes the message from the winning tab', () => {
		expect(peerPresenceOf('ada', idleFirst)).toEqual({
			status: 'available',
			message: 'at my desk',
		});
	});

	it('is offline with no tab and your own status for yourself', () => {
		expect(memberPresence('grace', 'me', 'available', idleFirst)).toBe('offline');
		expect(peerPresenceOf('grace', idleFirst)).toBeNull();
		expect(memberPresence('me', 'me', 'dnd', idleFirst)).toBe('dnd');
	});

	it('gives the contact list one answer per person', () => {
		const roster = [{ id: 'ada', name: 'Ada' }];
		const views = contactViews(roster as never, 'me', idleFirst);
		expect(views.map((v) => v.availability)).toEqual(['available']);
		expect(contactViews(roster as never, 'me', activeFirst)).toEqual(views);
	});

	it('carries the meeting activity of the winning tab to contacts', () => {
		const roster = [{ id: 'ada', name: 'Ada' }];
		const meeting = { ...peer(1, 'ada', 'busy'), activity: 'meeting' };
		const [view] = contactViews(roster as never, 'me', [meeting]);
		expect(view).toMatchObject({ availability: 'busy', activity: 'meeting' });
		const [set] = contactViews(roster as never, 'me', [meeting, peer(2, 'ada', 'busy')]);
		expect(set?.activity).toBeUndefined();
		const [free] = contactViews(roster as never, 'me', [peer(2, 'ada', 'available'), meeting]);
		expect(free).toMatchObject({ availability: 'busy', activity: 'meeting' });
	});
});
