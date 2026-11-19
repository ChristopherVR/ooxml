/**
 * collaboration-failure.ts: why a PowerPoint collaboration session could not
 * start, as one typed reason every binding renders the same way.
 *
 * The session controller catches every start error. The places that know the
 * cause throw a {@link CollaborationStartError} carrying its
 * {@link CollaborationFailure}; anything else is `unknown`. Each reason maps
 * to one translation key ({@link collaborationFailureMessageKey}), so the
 * status pill, the Share dialog and custom shells all say the same thing.
 *
 * @module render/collaboration-failure
 */
import type { TranslationKey } from '../i18n';
import type { ConnectionStatus } from './collaboration-presence';

/** The reasons a session start fails that the code can actually detect. */
export type CollaborationFailure =
	/** The room id is not 1 to 128 letters, digits, hyphens or underscores. */
	| 'invalid-room'
	/** A peer-to-peer (webrtc) session has no valid secret from its invitation link. */
	| 'missing-secret'
	/** A password-protected peer-to-peer room needs Web Crypto, which an insecure page lacks. */
	| 'insecure-page'
	/** `yjs`, `y-webrtc`, `y-websocket` or the collaboration runtime failed to load. */
	| 'module-unavailable'
	/** An https page refused an insecure `ws://` server. */
	| 'insecure-server'
	/** The websocket server did not connect before the timeout. */
	| 'server-unreachable'
	/** A same-browser room was asked for in a browser without BroadcastChannel. */
	| 'no-broadcast'
	/** Anything else, including a host provider that threw. */
	| 'unknown';

const MESSAGE_KEYS: Record<CollaborationFailure, TranslationKey> = {
	'invalid-room': 'pptx.collaboration.failure.invalidRoom',
	'missing-secret': 'pptx.collaboration.failure.missingSecret',
	'insecure-page': 'pptx.collaboration.failure.insecurePage',
	'module-unavailable': 'pptx.collaboration.failure.moduleUnavailable',
	'insecure-server': 'pptx.collaboration.failure.insecureServer',
	'server-unreachable': 'pptx.collaboration.failure.serverUnreachable',
	'no-broadcast': 'pptx.collaboration.failure.noBroadcast',
	unknown: 'pptx.collaboration.failure.unknown',
};

/** A start error whose cause is known. */
export class CollaborationStartError extends Error {
	readonly failure: CollaborationFailure;

	constructor(failure: CollaborationFailure, message: string, options?: { cause?: unknown }) {
		super(message, options);
		this.name = 'CollaborationStartError';
		this.failure = failure;
	}
}

function isCollaborationFailure(value: unknown): value is CollaborationFailure {
	return typeof value === 'string' && Object.hasOwn(MESSAGE_KEYS, value);
}

/** The reason behind a caught start error (`unknown` when nothing classified it). */
export function collaborationFailureOf(error: unknown): CollaborationFailure {
	// Structural, so an error thrown by a second copy of this module still classifies.
	const failure = (error as { failure?: unknown } | null)?.failure;
	return isCollaborationFailure(failure) ? failure : 'unknown';
}

/**
 * The translation key that explains an errored session, or `null` while the
 * status is not `error`. A borrowed host session that errors carries no
 * reason; it gets the transport-neutral `unknown` text.
 */
export function collaborationFailureMessageKey(
	status: ConnectionStatus,
	failure: CollaborationFailure | null,
): TranslationKey | null {
	return status === 'error' ? MESSAGE_KEYS[failure ?? 'unknown'] : null;
}

/** Run a lazy module loader, classifying a rejected import as `module-unavailable`. */
export async function loadCollaborationModule<T>(name: string, load: () => Promise<T>): Promise<T> {
	try {
		return await load();
	} catch (cause) {
		throw new CollaborationStartError('module-unavailable', `Could not load ${name}`, { cause });
	}
}

/**
 * Throw `insecure-page` when Web Crypto is missing. y-webrtc derives the room
 * key from its password with `crypto.subtle`, which only secure pages (https or
 * localhost) have; without it the provider constructor throws a bare TypeError.
 */
export function requirePeerToPeerCrypto(
	subtle: unknown = (globalThis as { crypto?: { subtle?: unknown } }).crypto?.subtle,
): void {
	if (!subtle) {
		throw new CollaborationStartError(
			'insecure-page',
			'Peer-to-peer collaboration needs Web Crypto, which only secure (https or localhost) pages have',
		);
	}
}
