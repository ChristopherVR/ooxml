// One WebRTC peer connection using the "perfect negotiation" pattern: the side with the smaller id
// is polite and yields on an offer collision, so both sides can add tracks at any time without a
// glare deadlock. Audio and video travel on two permanent transceivers whose tracks are swapped
// with `replaceTrack`, so muting, turning the camera off and sharing the screen never renegotiate.
// The connection and stream types are structural (no DOM lib), so it runs against fakes in tests.
// New code.
import type { IceCandidateInit, SessionDescription } from './signaling.js';

export interface TrackLike {
	kind: string;
	enabled: boolean;
	stop: () => void;
}
export interface StreamLike {
	getTracks: () => TrackLike[];
	addTrack: (track: TrackLike) => void;
}
export interface SenderLike {
	replaceTrack: (track: TrackLike | null) => Promise<void>;
}
export interface TransceiverLike {
	sender: SenderLike;
}
export interface PeerConnectionLike {
	signalingState: string;
	connectionState: string;
	onnegotiationneeded: (() => void) | null;
	onicecandidate: ((event: { candidate: IceCandidateInit | null }) => void) | null;
	ontrack: ((event: { track: TrackLike; streams: readonly StreamLike[] }) => void) | null;
	onconnectionstatechange: (() => void) | null;
	addTransceiver: (kind: 'audio' | 'video', init?: { direction: string }) => TransceiverLike;
	setLocalDescription: (description?: SessionDescription) => Promise<void>;
	setRemoteDescription: (description: SessionDescription) => Promise<void>;
	addIceCandidate: (candidate: IceCandidateInit) => Promise<void>;
	/** The description set by the last `setLocalDescription()`. */
	localDescription: { type: string; sdp?: string } | null;
	restartIce?: () => void;
	close: () => void;
}

export interface IceServer {
	urls: string | string[];
	username?: string;
	credential?: string;
}

export type PeerConnectionFactory = (config: {
	iceServers: IceServer[];
	iceTransportPolicy?: 'all' | 'relay';
}) => PeerConnectionLike;
export type StreamFactory = () => StreamLike;

export interface PeerCallbacks {
	sendDescription: (description: SessionDescription) => void;
	sendCandidate: (candidate: IceCandidateInit) => void;
	onStream: (stream: StreamLike) => void;
	onConnection: (state: string) => void;
	onError: (error: Error) => void;
}

export interface Peer {
	setTracks: (tracks: { audio?: TrackLike | null; video?: TrackLike | null }) => Promise<void>;
	receiveDescription: (description: SessionDescription) => Promise<void>;
	receiveCandidate: (candidate: IceCandidateInit) => Promise<void>;
	close: () => void;
}

export function createPeer(
	selfId: string,
	peerId: string,
	factory: PeerConnectionFactory,
	createStream: StreamFactory,
	config: { iceServers: IceServer[]; iceTransportPolicy?: 'all' | 'relay' },
	tracks: { audio?: TrackLike | null; video?: TrackLike | null },
	callbacks: PeerCallbacks,
): Peer {
	const pc = factory(config);
	const polite = selfId < peerId;
	let makingOffer = false;
	let ignoreOffer = false;
	let closed = false;
	const fail = (error: unknown): void => {
		if (!closed) callbacks.onError(error instanceof Error ? error : new Error(String(error)));
	};

	const audio = pc.addTransceiver('audio', { direction: 'sendrecv' });
	const video = pc.addTransceiver('video', { direction: 'sendrecv' });
	if (tracks.audio) void audio.sender.replaceTrack(tracks.audio).catch(fail);
	if (tracks.video) void video.sender.replaceTrack(tracks.video).catch(fail);

	let remote: StreamLike | null = null;
	pc.ontrack = ({ track, streams }) => {
		const stream = streams[0] ?? remote ?? createStream();
		if (!streams[0] && !stream.getTracks().includes(track)) stream.addTrack(track);
		if (stream !== remote) {
			remote = stream;
			callbacks.onStream(stream);
		}
	};
	pc.onicecandidate = ({ candidate }) => {
		if (candidate) callbacks.sendCandidate(candidate);
	};
	pc.onconnectionstatechange = () => {
		callbacks.onConnection(pc.connectionState);
		if (pc.connectionState === 'failed') pc.restartIce?.();
	};
	pc.onnegotiationneeded = () => {
		void (async () => {
			try {
				makingOffer = true;
				await pc.setLocalDescription();
				const d = pc.localDescription;
				if (d && (d.type === 'offer' || d.type === 'answer') && d.sdp)
					callbacks.sendDescription({ type: d.type, sdp: d.sdp });
			} catch (error) {
				fail(error);
			} finally {
				makingOffer = false;
			}
		})();
	};

	return {
		async setTracks(next) {
			if ('audio' in next) await audio.sender.replaceTrack(next.audio ?? null);
			if ('video' in next) await video.sender.replaceTrack(next.video ?? null);
		},
		async receiveDescription(description) {
			try {
				const collision =
					description.type === 'offer' && (makingOffer || pc.signalingState !== 'stable');
				ignoreOffer = !polite && collision;
				if (ignoreOffer) return;
				await pc.setRemoteDescription(description);
				if (description.type === 'offer') {
					await pc.setLocalDescription();
					const d = pc.localDescription;
					if (d?.sdp && (d.type === 'answer' || d.type === 'offer'))
						callbacks.sendDescription({ type: d.type, sdp: d.sdp });
				}
			} catch (error) {
				fail(error);
			}
		},
		async receiveCandidate(candidate) {
			try {
				await pc.addIceCandidate(candidate);
			} catch (error) {
				if (!ignoreOffer) fail(error);
			}
		},
		close() {
			closed = true;
			pc.onnegotiationneeded = pc.onicecandidate = pc.ontrack = pc.onconnectionstatechange = null;
			pc.close();
		},
	};
}
