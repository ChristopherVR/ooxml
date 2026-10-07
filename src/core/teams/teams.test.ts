import * as Y from 'yjs';
import { createCallSession } from './call';
import { createChatStore } from './chat';
import { detectOfficeKind, sanitizeAttachment, sanitizeMessageText } from './model';
import type { PeerConnectionLike, StreamLike, TrackLike } from './peer';
import { endpointUrl, parseServerConfig } from './server-config';
import { createMemorySignalingHub, parseSignal } from './signaling';

const sync = (a: Y.Doc, b: Y.Doc): void => {
	Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
	Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
};
const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

describe('model', () => {
	it('detects Office kinds from the extension before the MIME type', () => {
		expect(detectOfficeKind('Budget.XLSX')).toBe('xlsx');
		expect(detectOfficeKind('plan.vsdx')).toBe('vsdx');
		expect(
			detectOfficeKind(
				'x',
				'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			),
		).toBe('pptx');
		expect(detectOfficeKind('notes.txt')).toBe('other');
	});
	it('strips markup from messages and drops unsafe attachment links', () => {
		expect(sanitizeMessageText('hi <img src=x onerror=alert(1)> there')).not.toContain('<');
		expect(sanitizeAttachment({ name: 'a.docx', url: 'javascript:alert(1)' })?.url).toBeUndefined();
		expect(sanitizeAttachment({ name: 'a.docx', url: 'https://h/a.docx' })?.url).toBe(
			'https://h/a.docx',
		);
		expect(sanitizeAttachment({ name: 'a.docx', url: '//evil/a' })?.url).toBeUndefined();
	});
});

describe('chat store', () => {
	it('converges channels, messages and concurrent reactions across two docs', () => {
		const a = new Y.Doc();
		const b = new Y.Doc();
		const alice = createChatStore(a, { id: 'alice', name: 'Alice' });
		const bob = createChatStore(b, { id: 'bob', name: 'Bob' });
		const ch = alice.createChannel({ id: 'general', name: 'General' });
		expect(ch?.id).toBe('general');
		sync(a, b);
		const m = alice.post('general', { text: 'hello <b>team</b>' });
		sync(a, b);
		bob.react(m!.id, '👍');
		alice.react(m!.id, '👍');
		sync(a, b);
		for (const store of [alice, bob]) {
			const [msg] = store.messages('general');
			expect(msg?.text).toBe('hello team');
			expect(msg?.reactions['👍']?.sort()).toEqual(['alice', 'bob']);
		}
	});
	it('only lets the author edit or delete', () => {
		const a = new Y.Doc();
		const b = new Y.Doc();
		const alice = createChatStore(a, { id: 'alice', name: 'Alice' });
		const bob = createChatStore(b, { id: 'bob', name: 'Bob' });
		alice.createChannel({ id: 'general', name: 'General' });
		const m = alice.post('general', { text: 'one' })!;
		sync(a, b);
		expect(bob.edit('general', m.id, 'hacked')).toBe(false);
		expect(bob.remove('general', m.id)).toBe(false);
		expect(alice.edit('general', m.id, 'two')).toBe(true);
		expect(alice.remove('general', m.id)).toBe(true);
		expect(alice.messages('general')[0]).toMatchObject({ deleted: true, text: '' });
	});
	it('refuses empty posts and unknown channels', () => {
		const store = createChatStore(new Y.Doc(), { id: 'a', name: 'A' });
		expect(store.post('nope', { text: 'x' })).toBeNull();
		store.createChannel({ id: 'c', name: 'C' });
		expect(store.post('c', { text: '   ' })).toBeNull();
	});
	it('notifies observers once per burst', async () => {
		const doc = new Y.Doc();
		const store = createChatStore(doc, { id: 'a', name: 'A' });
		let n = 0;
		store.observe(() => n++);
		store.createChannel({ id: 'c', name: 'C' });
		store.post('c', { text: 'x' });
		await tick();
		expect(n).toBe(1);
	});
});

describe('signaling', () => {
	it('rejects malformed, oversized and mistyped signals', () => {
		expect(parseSignal('not json')).toBeNull();
		expect(
			parseSignal({
				type: 'description',
				from: 'a',
				to: 'b',
				description: { type: 'rollback', sdp: 'x' },
			}),
		).toBeNull();
		expect(
			parseSignal({ type: 'description', from: 'a', description: { type: 'offer', sdp: 'x' } }),
		).toBeNull();
		expect(parseSignal({ type: 'hello', from: 'a' })).toBeNull();
		expect(parseSignal({ type: 'bye', from: 'a' })).toEqual({ type: 'bye', from: 'a' });
	});
});

describe('server config', () => {
	it('validates URLs and TURN credentials and reports issues', () => {
		const { config, issues } = parseServerConfig({
			mode: 'server',
			syncUrl: 'https://not-a-socket',
			signalingUrl: 'wss://h.example/signal/',
			iceServers: [{ urls: 'turn:t.example' }, { urls: ['stun:s.example'] }],
		});
		expect(config.syncUrl).toBeUndefined();
		expect(config.signalingUrl).toBe('wss://h.example/signal');
		expect(config.iceServers).toEqual([{ urls: 'stun:s.example' }]);
		expect(issues.length).toBe(2);
	});
	it('falls back to local without endpoints and builds endpoint URLs', () => {
		expect(parseServerConfig({ mode: 'server' }).config.mode).toBe('local');
		expect(endpointUrl('wss://h/sync/', 'room 1', 't/k')).toBe('wss://h/sync/room%201?token=t%2Fk');
	});
});

// A fake connection pair: whatever one side sets as local description the other can accept.
class FakePc implements PeerConnectionLike {
	signalingState = 'stable';
	connectionState = 'new';
	localDescription: { type: string; sdp?: string } | null = null;
	onnegotiationneeded: (() => void) | null = null;
	onicecandidate: PeerConnectionLike['onicecandidate'] = null;
	ontrack: PeerConnectionLike['ontrack'] = null;
	onconnectionstatechange: (() => void) | null = null;
	closed = false;
	senders: (TrackLike | null)[] = [];
	private n = 0;
	addTransceiver(): ReturnType<PeerConnectionLike['addTransceiver']> {
		const i = this.senders.push(null) - 1;
		if (++this.n === 2) queueMicrotask(() => this.onnegotiationneeded?.());
		return { sender: { replaceTrack: async (t) => void (this.senders[i] = t) } };
	}
	async setLocalDescription(): Promise<void> {
		this.localDescription = {
			type: this.signalingState === 'have-remote-offer' ? 'answer' : 'offer',
			sdp: 'v=0',
		};
		this.signalingState = this.localDescription.type === 'offer' ? 'have-local-offer' : 'stable';
	}
	async setRemoteDescription(d: { type: string }): Promise<void> {
		this.signalingState = d.type === 'offer' ? 'have-remote-offer' : 'stable';
		if (d.type === 'answer') this.connectionState = 'connected';
		else
			queueMicrotask(
				() => ((this.connectionState = 'connected'), this.onconnectionstatechange?.()),
			);
	}
	async addIceCandidate(): Promise<void> {}
	close(): void {
		this.closed = true;
	}
}
const track = (kind: string): TrackLike => ({ kind, enabled: true, stop() {} });
const stream = (): StreamLike => {
	const tracks: TrackLike[] = [];
	return { getTracks: () => tracks, addTrack: (t) => void tracks.push(t) };
};

describe('call session', () => {
	const make = (hub: ReturnType<typeof createMemorySignalingHub>, id: string, pcs: FakePc[]) =>
		createCallSession({
			roomId: 'r1',
			user: { id, name: id.toUpperCase() },
			signaling: hub.createChannel('r1'),
			createPeerConnection: () => {
				const pc = new FakePc();
				pcs.push(pc);
				return pc;
			},
			createStream: stream,
			mediaDevices: {
				getUserMedia: async () => ({
					getTracks: () => [track('audio'), track('video')],
					addTrack() {},
				}),
			},
		});

	it('discovers peers, negotiates and tracks media state', async () => {
		const hub = createMemorySignalingHub();
		const pa: FakePc[] = [];
		const pb: FakePc[] = [];
		const a = make(hub, 'a', pa);
		const b = make(hub, 'b', pb);
		await a.join();
		await tick();
		await b.join();
		await tick();
		await tick();
		expect(a.participants().map((p) => p.id)).toEqual(['a', 'b']);
		expect(b.participants().map((p) => p.name)).toEqual(['B', 'A']);
		expect(pa).toHaveLength(1);
		expect(pb).toHaveLength(1);
		expect(a.participants()[1]?.connection).toBe('connected');
		a.toggleMic();
		await tick();
		expect(b.participants().find((p) => p.id === 'a')?.state.audio).toBe(false);
		a.toggleHand();
		await tick();
		expect(b.participants().find((p) => p.id === 'a')?.state.hand).toBe(true);
		b.leave();
		await tick();
		expect(a.participants().map((p) => p.id)).toEqual(['a']);
		expect(pb[0]?.closed).toBe(true);
	});

	it('refuses peers beyond the mesh ceiling', async () => {
		const hub = createMemorySignalingHub();
		const errors: string[] = [];
		const a = createCallSession({
			roomId: 'r1',
			user: { id: 'a', name: 'A' },
			signaling: hub.createChannel('r1'),
			createPeerConnection: () => new FakePc(),
			createStream: stream,
			maxParticipants: 2,
		});
		a.on('error', (e) => errors.push(e.message));
		await a.join();
		await tick();
		for (const id of ['b', 'c']) {
			const ch = hub.createChannel('r1');
			ch.connect({ open() {}, close() {}, message() {}, error() {} });
			ch.send({
				type: 'hello',
				from: id,
				name: id,
				state: { audio: false, video: false, screen: false, hand: false },
			});
		}
		await tick();
		expect(a.participants()).toHaveLength(2);
		expect(errors.some((e) => e.includes('full'))).toBe(true);
	});
});
