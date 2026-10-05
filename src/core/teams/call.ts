// A call: a full-mesh set of peer connections over a `SignalingChannel`, with the local media state
// (microphone, camera, screen share, raised hand). Mesh is the honest scope for a bring-your-own
// stack: it needs no media server but scales to roughly a dozen people; an SFU would plug in behind
// the same `CallSession` surface. Media capture, the peer connection constructor and the stream
// constructor are injectable, so the session is testable without a browser. DOM-free. New code.
import { Emitter } from '../collab/emitter.js';
import {
	type IceServer,
	type Peer,
	type PeerConnectionFactory,
	type StreamFactory,
	type StreamLike,
	type TrackLike,
	createPeer,
} from './peer.js';
import type { MediaState, Signal, SignalingChannel } from './signaling.js';

export interface MediaDevicesLike {
	getUserMedia: (constraints: { audio?: boolean; video?: boolean }) => Promise<StreamLike>;
	getDisplayMedia?: (constraints?: { audio?: boolean; video?: boolean }) => Promise<StreamLike>;
}

export type PeerConnectionState =
	| 'new'
	| 'connecting'
	| 'connected'
	| 'disconnected'
	| 'failed'
	| 'closed';

export interface CallParticipant {
	id: string;
	name: string;
	self: boolean;
	state: MediaState;
	connection: PeerConnectionState;
	stream: StreamLike | null;
}

export type CallStatus = 'idle' | 'joining' | 'connected' | 'left';

export interface CallEvents {
	participants: CallParticipant[];
	status: CallStatus;
	error: Error;
}

export interface CallOptions {
	roomId: string;
	user: { id: string; name: string };
	signaling: SignalingChannel;
	iceServers?: IceServer[];
	/** `relay` forces TURN (hides peer IPs, needs a TURN server). Default `all`. */
	iceTransportPolicy?: 'all' | 'relay';
	createPeerConnection?: PeerConnectionFactory;
	createStream?: StreamFactory;
	mediaDevices?: MediaDevicesLike;
	/** Mesh ceiling, counting yourself. Default 12. */
	maxParticipants?: number;
}

export interface CallSession {
	readonly roomId: string;
	readonly status: CallStatus;
	participants: () => CallParticipant[];
	join: (options?: { audio?: boolean; video?: boolean }) => Promise<void>;
	leave: () => void;
	toggleMic: () => void;
	toggleCamera: () => Promise<void>;
	startScreenShare: () => Promise<void>;
	stopScreenShare: () => Promise<void>;
	toggleHand: () => void;
	on: <K extends keyof CallEvents>(
		event: K,
		listener: (payload: CallEvents[K]) => void,
	) => () => void;
}

const EMPTY: MediaState = { audio: false, video: false, screen: false, hand: false };
const asConnection = (s: string): PeerConnectionState =>
	s === 'connected' ||
	s === 'connecting' ||
	s === 'disconnected' ||
	s === 'failed' ||
	s === 'closed'
		? s
		: 'new';

interface Remote {
	name: string;
	state: MediaState;
	connection: PeerConnectionState;
	stream: StreamLike | null;
	peer: Peer;
}

function defaults(): {
	pc: PeerConnectionFactory;
	stream: StreamFactory;
	media?: MediaDevicesLike;
} {
	const g = globalThis as Record<string, unknown>;
	const PC = g.RTCPeerConnection as (new (c: unknown) => unknown) | undefined;
	const MS = g.MediaStream as (new () => unknown) | undefined;
	const nav = g.navigator as { mediaDevices?: MediaDevicesLike } | undefined;
	return {
		pc: (config) => {
			if (!PC) throw new Error('RTCPeerConnection is not available in this environment');
			return new PC(config) as ReturnType<PeerConnectionFactory>;
		},
		stream: () => {
			if (!MS) throw new Error('MediaStream is not available in this environment');
			return new MS() as StreamLike;
		},
		...(nav?.mediaDevices ? { media: nav.mediaDevices } : {}),
	};
}

export function createCallSession(options: CallOptions): CallSession {
	const env = defaults();
	const makePc = options.createPeerConnection ?? env.pc;
	const makeStream = options.createStream ?? env.stream;
	const media = options.mediaDevices ?? env.media;
	const maxParticipants = options.maxParticipants ?? 12;
	const events = new Emitter<CallEvents>();
	const self = options.user;
	const remotes = new Map<string, Remote>();
	const local: { audio: TrackLike | null; camera: TrackLike | null; screen: TrackLike | null } = {
		audio: null,
		camera: null,
		screen: null,
	};
	let state: MediaState = { ...EMPTY };
	let status: CallStatus = 'idle';
	let selfStream: StreamLike | null = null;

	const setStatus = (next: CallStatus): void => {
		status = next;
		events.emit('status', next);
	};
	const videoTrack = (): TrackLike | null => local.screen ?? local.camera;
	const rebuildSelfStream = (): void => {
		const tracks = [local.audio, videoTrack()].filter((t): t is TrackLike => t !== null);
		if (tracks.length === 0) return void (selfStream = null);
		const stream = makeStream();
		for (const t of tracks) stream.addTrack(t);
		selfStream = stream;
	};
	const list = (): CallParticipant[] => [
		{
			id: self.id,
			name: self.name,
			self: true,
			state,
			connection: 'connected',
			stream: selfStream,
		},
		...[...remotes].map(([id, r]) => ({
			id,
			name: r.name,
			self: false,
			state: r.state,
			connection: r.connection,
			stream: r.stream,
		})),
	];
	const publish = (): void => events.emit('participants', list());
	const send = (signal: Signal): void => options.signaling.send(signal);
	const announceState = (): void => send({ type: 'state', from: self.id, state });
	const fail = (error: unknown): void =>
		events.emit('error', error instanceof Error ? error : new Error(String(error)));

	const ensurePeer = (id: string, name: string, remoteState: MediaState): Remote | null => {
		const existing = remotes.get(id);
		if (existing) {
			existing.name = name;
			existing.state = remoteState;
			return existing;
		}
		if (remotes.size + 1 >= maxParticipants) {
			fail(new Error(`Call is full (${maxParticipants} participants)`));
			return null;
		}
		const config = {
			iceServers: options.iceServers ?? [],
			...(options.iceTransportPolicy ? { iceTransportPolicy: options.iceTransportPolicy } : {}),
		};
		const remote: Remote = {
			name,
			state: remoteState,
			connection: 'new',
			stream: null,
			peer: createPeer(
				self.id,
				id,
				makePc,
				makeStream,
				config,
				{ audio: local.audio, video: videoTrack() },
				{
					sendDescription: (description) =>
						send({ type: 'description', from: self.id, to: id, description }),
					sendCandidate: (candidate) =>
						send({ type: 'candidate', from: self.id, to: id, candidate }),
					onStream: (stream) => {
						remote.stream = stream;
						publish();
					},
					onConnection: (s) => {
						remote.connection = asConnection(s);
						publish();
					},
					onError: fail,
				},
			),
		};
		remotes.set(id, remote);
		return remote;
	};
	const drop = (id: string): void => {
		const r = remotes.get(id);
		if (!r) return;
		r.peer.close();
		remotes.delete(id);
		publish();
	};

	const hello = (to?: string, reply = false): void =>
		send({
			type: 'hello',
			from: self.id,
			name: self.name,
			state,
			...(to ? { to } : {}),
			...(reply ? { reply } : {}),
		});

	const onSignal = (signal: Signal): void => {
		if (signal.from === self.id) return;
		switch (signal.type) {
			case 'hello':
				if (signal.to && signal.to !== self.id) return;
				if (ensurePeer(signal.from, signal.name, signal.state) && !signal.reply)
					hello(signal.from, true);
				publish();
				return;
			case 'bye':
				return drop(signal.from);
			case 'state': {
				const r = remotes.get(signal.from);
				if (r) r.state = signal.state;
				return publish();
			}
			case 'description':
				if (signal.to === self.id)
					void remotes.get(signal.from)?.peer.receiveDescription(signal.description);
				return;
			case 'candidate':
				if (signal.to === self.id)
					void remotes.get(signal.from)?.peer.receiveCandidate(signal.candidate);
				return;
		}
	};

	const setVideoEverywhere = async (): Promise<void> => {
		const track = videoTrack();
		await Promise.all([...remotes.values()].map((r) => r.peer.setTracks({ video: track })));
		rebuildSelfStream();
		publish();
	};
	const watchEnd = (track: TrackLike, onEnd: () => void): void => {
		(track as unknown as { onended: (() => void) | null }).onended = onEnd;
	};

	return {
		roomId: options.roomId,
		get status() {
			return status;
		},
		participants: list,
		on: (event, listener) => events.on(event, listener),
		async join(want = {}) {
			if (status === 'joining' || status === 'connected') return;
			setStatus('joining');
			const audio = want.audio ?? true;
			const video = want.video ?? true;
			if (media && (audio || video)) {
				let stream: StreamLike | null = null;
				for (const c of [
					{ audio, video },
					{ audio, video: false },
					{ audio: false, video },
				]) {
					if (!c.audio && !c.video) continue;
					try {
						stream = await media.getUserMedia(c);
						break;
					} catch (error) {
						if (c.audio && c.video) fail(error);
					}
				}
				for (const t of stream?.getTracks() ?? []) {
					if (t.kind === 'audio') local.audio = t;
					else if (t.kind === 'video') local.camera = t;
				}
			}
			state = { ...EMPTY, audio: Boolean(local.audio), video: Boolean(local.camera) };
			rebuildSelfStream();
			options.signaling.connect({
				open: () => {
					hello();
					setStatus('connected');
					publish();
				},
				close: () => {},
				message: onSignal,
				error: fail,
			});
			publish();
		},
		leave() {
			if (status === 'left' || status === 'idle') return;
			send({ type: 'bye', from: self.id });
			for (const id of [...remotes.keys()]) drop(id);
			for (const t of [local.audio, local.camera, local.screen]) t?.stop();
			local.audio = local.camera = local.screen = null;
			selfStream = null;
			options.signaling.disconnect();
			setStatus('left');
			publish();
		},
		toggleMic() {
			if (!local.audio) return;
			local.audio.enabled = !local.audio.enabled;
			state = { ...state, audio: local.audio.enabled };
			announceState();
			publish();
		},
		async toggleCamera() {
			if (local.camera) {
				local.camera.stop();
				local.camera = null;
			} else if (media) {
				try {
					const s = await media.getUserMedia({ video: true });
					local.camera = s.getTracks().find((t) => t.kind === 'video') ?? null;
				} catch (error) {
					return fail(error);
				}
			}
			state = { ...state, video: Boolean(local.camera) };
			await setVideoEverywhere();
			announceState();
		},
		async startScreenShare() {
			if (local.screen || !media?.getDisplayMedia) return;
			try {
				const s = await media.getDisplayMedia({ video: true });
				const track = s.getTracks().find((t) => t.kind === 'video') ?? null;
				if (!track) return;
				local.screen = track;
				watchEnd(track, () => void this.stopScreenShare());
			} catch (error) {
				return fail(error);
			}
			state = { ...state, screen: true };
			await setVideoEverywhere();
			announceState();
		},
		async stopScreenShare() {
			if (!local.screen) return;
			local.screen.stop();
			local.screen = null;
			state = { ...state, screen: false };
			await setVideoEverywhere();
			announceState();
		},
		toggleHand() {
			state = { ...state, hand: !state.hand };
			announceState();
			publish();
		},
	};
}
