import { describe, expect, it } from 'vitest';
import {
	canJoinShare,
	canStartShare,
	createShareSecret,
	invitationFromLocation,
	isValidShareSecret,
	joinShareOptions,
	parseShareInvitation,
	seedShareFields,
	shareInvitation,
	startShareOptions,
	suggestRoomId,
	webrtcRoomName,
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

	it('links a serverless session by room only', () => {
		expect(shareInvitation({ roomId: 'team', serverUrl: '' }, page)).toBe(
			'https://office.example/word/?room=team',
		);
	});

	it('carries a peer-to-peer secret only in the link fragment', () => {
		const secret = createShareSecret();
		const p2p = shareInvitation(
			{
				roomId: 'team',
				serverUrl: '',
				transport: 'webrtc',
				signaling: ['wss://a', 'wss://b'],
				secret,
			},
			page,
		);
		const url = new URL(p2p);
		expect(url.search).not.toContain(secret);
		expect(url.hash).toBe(`#key=${secret}`);
		expect(parseShareInvitation(p2p)).toEqual({
			roomId: 'team',
			serverUrl: '',
			transport: 'webrtc',
			signaling: ['wss://a', 'wss://b'],
			secret,
		});
	});

	it('refuses a peer-to-peer link without a valid secret', () => {
		const bare = 'https://office.example/word/?room=team&transport=webrtc';
		expect(parseShareInvitation(bare)).toBeNull();
		expect(parseShareInvitation(`${bare}#key=short`)).toBeNull();
		expect(parseShareInvitation(`${bare}&key=${createShareSecret()}`)).toBeNull();
		expect(canJoinShare({ invitation: bare, userName: 'Bo', serverUrl: '' })).toBe(false);
	});

	it('reads a broadcast follow link under its own parameter', () => {
		const secret = createShareSecret();
		const link = shareInvitation(
			{ roomId: 'show', serverUrl: '', transport: 'webrtc', secret },
			page,
			'broadcast',
		);
		expect(link).toBe(`https://office.example/word/?broadcast=show&transport=webrtc#key=${secret}`);
		expect(parseShareInvitation(link)).toBeNull();
		expect(parseShareInvitation(link, '', 'broadcast')).toMatchObject({ roomId: 'show', secret });
	});

	it('makes fresh 128-bit secrets and a signaling room that hides both names', () => {
		const secret = createShareSecret();
		expect(secret).toMatch(/^[A-Za-z0-9_-]{22}$/u);
		expect(isValidShareSecret(secret)).toBe(true);
		expect(createShareSecret()).not.toBe(secret);
		const room = webrtcRoomName(secret);
		expect(room).toMatch(/^[0-9a-f]{64}$/u);
		expect(room).toBe(webrtcRoomName(secret));
		expect(room).not.toContain(secret);
		expect(webrtcRoomName(createShareSecret())).not.toBe(room);
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
