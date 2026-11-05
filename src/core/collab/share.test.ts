import { describe, expect, it } from 'vitest';
import {
	canJoinShare,
	canStartShare,
	invitationFromLocation,
	joinShareOptions,
	parseShareInvitation,
	seedShareFields,
	shareInvitation,
	startShareOptions,
	suggestRoomId,
} from './share';

const page = { origin: 'https://office.example', pathname: '/word/' };

describe('share', () => {
	it('suggests valid room names', () => {
		const room = suggestRoomId('doc');
		expect(room).toMatch(/^doc-[a-z0-9]{8}$/u);
		expect(suggestRoomId('doc')).not.toBe(room);
	});

	it('starts only with a valid room and a name', () => {
		expect(canStartShare(seedShareFields({ roomId: 'team', userName: ' ' }))).toBe(false);
		expect(canStartShare(seedShareFields({ roomId: 'bad room', userName: 'Ann' }))).toBe(false);
		expect(startShareOptions(seedShareFields({ roomId: ' team ', userName: ' Ann ' }))).toEqual({
			roomId: 'team',
			user: { name: 'Ann' },
			intent: 'create',
		});
		expect(
			startShareOptions({ roomId: 'team', userName: 'Ann', serverUrl: ' wss://s.example ' }),
		).toEqual({
			roomId: 'team',
			serverUrl: 'wss://s.example',
			user: { name: 'Ann' },
			intent: 'create',
		});
	});

	it('round-trips an invitation link with its server', () => {
		const link = shareInvitation({ roomId: 'team', serverUrl: 'wss://s.example' }, page);
		expect(link).toBe('https://office.example/word/?room=team&server=wss%3A%2F%2Fs.example');
		expect(parseShareInvitation(link)).toEqual({ roomId: 'team', serverUrl: 'wss://s.example' });
		expect(joinShareOptions({ invitation: link, userName: 'Bo', serverUrl: '' })).toEqual({
			roomId: 'team',
			serverUrl: 'wss://s.example',
			user: { name: 'Bo' },
			intent: 'join',
		});
	});

	it('links a serverless session by room only, and keeps peer-to-peer parameters', () => {
		expect(shareInvitation({ roomId: 'team', serverUrl: '' }, page)).toBe(
			'https://office.example/word/?room=team',
		);
		const p2p = shareInvitation(
			{ roomId: 'team', serverUrl: '', transport: 'webrtc', signaling: ['wss://a', 'wss://b'] },
			page,
		);
		expect(parseShareInvitation(p2p)).toEqual({
			roomId: 'team',
			serverUrl: '',
			transport: 'webrtc',
			signaling: ['wss://a', 'wss://b'],
		});
	});

	it('falls back to the room code without a usable location', () => {
		expect(shareInvitation({ roomId: 'team', serverUrl: 'wss://s' })).toBe('team');
		expect(
			shareInvitation({ roomId: 'team', serverUrl: '' }, { origin: 'null', pathname: '/' }),
		).toBe('team');
	});

	it('joins a bare room code through the fallback server', () => {
		expect(parseShareInvitation(' team ', 'wss://mine')).toEqual({
			roomId: 'team',
			serverUrl: 'wss://mine',
		});
		expect(canJoinShare({ invitation: 'team', userName: 'Bo', serverUrl: '' })).toBe(true);
		expect(canJoinShare({ invitation: 'team', userName: '', serverUrl: '' })).toBe(false);
		expect(
			canJoinShare({
				invitation: 'https://x.example/?room=bad%20room',
				userName: 'Bo',
				serverUrl: '',
			}),
		).toBe(false);
		expect(parseShareInvitation('')).toBeNull();
	});

	it('reads the invitation a page was opened with', () => {
		expect(invitationFromLocation('https://x.example/?room=team')).toBe(
			'https://x.example/?room=team',
		);
		expect(invitationFromLocation('https://x.example/')).toBeNull();
		expect(invitationFromLocation('not a url')).toBeNull();
	});
});
