/**
 * The collaboration session lifecycle every PowerPoint binding adapts.
 *
 * Each binding used to carry its own copy of "join a room": validate the room
 * id, open a provider or same-browser room when the config needs one
 * (`openOwnedCollaboration`), resolve the transport, refuse a ws:// socket
 * from an https page, build the Y.Doc and provider, drive the connection
 * status (immediately connected for WebRTC, a connect timeout for websocket, a
 * drop back to `disconnected` that re-gates writes), open the first-write gate
 * on the provider's sync, leave the room on `pagehide`, and tear it all down
 * again. This controller is that lifecycle, once.
 *
 * What stays in the binding is the document and presence mapping: it hears
 * {@link CollaborationSessionControllerDeps.onOpen} with the live session and
 * {@link CollaborationSessionControllerDeps.onClose} before it is disposed.
 * A borrowed host session (`externalSession`) has no transport: the binding's
 * readiness logic reports its status through {@link
 * CollaborationSessionController.setStatus} and opens the gate itself.
 *
 * @module render/controller/collaboration-session-controller
 */

import type { CollaborationConfig, CollaborationTransport } from '../../types';
import { resolveTransportForServerUrl } from '../broadcast-helpers';
import { resolveCollaborationConfig } from '../collaboration-options';
import type { CollaborationInput } from '../collaboration-options';
import { needsOwnedSession, openOwnedCollaboration } from '../collaboration-owned-session';
import type { ConnectionStatus } from '../collaboration-presence';
import {
	CONNECTION_TIMEOUT_MS,
	isMixedContentBlocked,
	validateRoomId,
} from '../collaboration-presence';
import { createSyncGate } from '../collaboration-sync-gate';
import type { SyncGate } from '../collaboration-sync-gate';
import { registerCollaborationTeardown } from '../collaboration-teardown';
import type { CollabYjsModules, CollabYSession } from '../collaboration-yjs-session';
import { openCollaborationYSession } from '../collaboration-yjs-session';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** What the chrome and the document layer read from the session. */
export interface CollaborationSessionState {
	readonly status: ConnectionStatus;
	/** A session is open (its document and presence are live). */
	readonly active: boolean;
	/** The config the session was started with (its identity), while started. */
	readonly config: CollaborationConfig | null;
	/** The effective config once a provider or same-browser room prepared it. */
	readonly prepared: CollaborationConfig | null;
	readonly transport: CollaborationTransport | null;
	/** The local awareness client id, while active. */
	readonly clientId: number | null;
	/** The first-write gate is open (only tracked for the controller's own gate). */
	readonly synced: boolean;
}

/** Timer seam; defaults to the global timers. */
export interface CollaborationSessionTimers {
	setTimeout(handler: () => void, ms: number): unknown;
	clearTimeout(id: unknown): void;
}

export interface CollaborationSessionControllerDeps {
	/** The binding's lazy `yjs` / `y-websocket` / `y-webrtc` loaders. */
	modules?: CollabYjsModules;
	/** Open the session; defaults to `openCollaborationYSession` with {@link modules}. */
	openSession?(
		config: CollaborationConfig,
		transport: CollaborationTransport,
	): Promise<CollabYSession>;
	/**
	 * The binding's own first-write gate. Without one the controller keeps its
	 * own (see {@link CollaborationSessionController.gate}) and reports
	 * `synced` in its state.
	 */
	gate?: SyncGate;
	/** The controller's own gate opened: perform the deferred first write. */
	onSynced?(): void;
	/** A session is starting with `config` (again once a room prepared it). */
	onBegin?(config: CollaborationConfig): void;
	/** The session is live: attach the document and presence mapping. */
	onOpen?(
		session: CollabYSession,
		config: CollaborationConfig,
		transport: CollaborationTransport,
	): void;
	/** The session is ending (also on every restart): detach before it is disposed. */
	onClose?(): void;
	/** Leave the room on `pagehide` and rejoin from bfcache (default true). */
	registerTeardown?: boolean;
	connectTimeoutMs?: number;
	timers?: CollaborationSessionTimers;
}

export interface CollaborationSessionController extends PptxController<CollaborationSessionState> {
	/** Join (or re-join) a room; any running session is stopped first. */
	start(input: CollaborationInput): Promise<void>;
	/** Leave the room and release everything this controller opened. */
	stop(): void;
	/** Start again with the last input, after an error or a drop. */
	retry(): Promise<void>;
	/** Report the status of a borrowed host session (its readiness owns it). */
	setStatus(status: ConnectionStatus): void;
	/** The live session, while active. */
	getSession(): CollabYSession | null;
	/** The first-write gate in use (the binding's, or the controller's own). */
	readonly gate: SyncGate;
	destroy(): void;
}

const INITIAL: CollaborationSessionState = {
	status: 'disconnected',
	active: false,
	config: null,
	prepared: null,
	transport: null,
	clientId: null,
	synced: false,
};

const defaultTimers: CollaborationSessionTimers = {
	setTimeout: (handler, ms) => setTimeout(handler, ms),
	clearTimeout: (id) => clearTimeout(id as ReturnType<typeof setTimeout>),
};

/** Build the session lifecycle controller (see the module note). */
export function createCollaborationSessionController(
	deps: CollaborationSessionControllerDeps,
): CollaborationSessionController {
	const timers = deps.timers ?? defaultTimers;
	const store = createControllerStore<CollaborationSessionState>(INITIAL);
	const patch = (next: Partial<CollaborationSessionState>): void => {
		const current = store.getSnapshot();
		const changed = (Object.keys(next) as (keyof CollaborationSessionState)[]).some(
			(key) => next[key] !== current[key],
		);
		if (changed) {
			store.set({ ...current, ...next });
		}
	};
	const openSession =
		deps.openSession ??
		((config: CollaborationConfig, transport: CollaborationTransport) => {
			if (!deps.modules) {
				return Promise.reject(new Error('No collaboration modules were provided'));
			}
			return openCollaborationYSession(config, transport, deps.modules);
		});

	const ownGate = deps.gate
		? null
		: createSyncGate(() => {
				patch({ synced: true });
				deps.onSynced?.();
			});
	const gate: SyncGate = deps.gate ?? {
		isOpen: () => ownGate!.isOpen(),
		arm: () => ownGate!.arm(),
		open: () => ownGate!.open(),
		reset: () => {
			ownGate!.reset();
			patch({ synced: false });
		},
	};

	let generation = 0;
	let session: CollabYSession | null = null;
	let releaseOwned: (() => void) | null = null;
	let connectTimer: unknown = null;
	let lastInput: CollaborationInput | null = null;
	let destroyed = false;

	const cancelConnectTimer = (): void => {
		if (connectTimer !== null) {
			timers.clearTimeout(connectTimer);
			connectTimer = null;
		}
	};
	const setStatus = (status: ConnectionStatus): void => patch({ status });

	const stop = (): void => {
		generation += 1;
		cancelConnectTimer();
		gate.reset();
		deps.onClose?.();
		session?.dispose();
		session = null;
		releaseOwned?.();
		releaseOwned = null;
		patch(INITIAL);
	};

	const fail = (token: number): void => {
		if (token === generation) {
			stop();
			setStatus('error');
		}
	};

	/** Status and gate wiring for an owned transport. */
	const wireProvider = (live: CollabYSession, transport: CollaborationTransport): void => {
		const provider = live.provider;
		if (!provider) {
			return;
		}
		provider.onSynced(() => gate.open());
		if (provider.syncedNow) {
			gate.open();
		} else {
			gate.arm();
		}
		provider.onStatus((connected) => {
			if (session !== live) {
				return;
			}
			if (connected) {
				cancelConnectTimer();
				setStatus('connected');
			} else {
				// Re-gate writes after a drop, so a peer that rejoins cannot clobber
				// the room with a stale local document.
				setStatus('disconnected');
				gate.reset();
				gate.arm();
			}
		});
		if (transport === 'webrtc' || provider.connectedNow) {
			// Same-browser WebRTC peers meet over BroadcastChannel at once.
			setStatus('connected');
			return;
		}
		const token = generation;
		connectTimer = timers.setTimeout(() => {
			connectTimer = null;
			if (token === generation && store.getSnapshot().status !== 'connected') {
				fail(token);
			}
		}, deps.connectTimeoutMs ?? CONNECTION_TIMEOUT_MS);
	};

	const start = async (input: CollaborationInput): Promise<void> => {
		if (destroyed) {
			return;
		}
		const requested = resolveCollaborationConfig(input);
		stop();
		if (!requested) {
			return;
		}
		const token = generation;
		lastInput = input;
		patch({ config: requested, prepared: requested });
		deps.onBegin?.(requested);
		let config = requested;
		try {
			validateRoomId(requested.roomId);
			if (needsOwnedSession(requested)) {
				setStatus('connecting');
				const owned = openOwnedCollaboration(requested);
				releaseOwned = owned.release;
				const prepared = await owned.ready;
				if (token !== generation || !prepared) {
					return;
				}
				config = prepared;
				patch({ prepared });
				deps.onBegin?.(prepared);
			}
			const transport = config.transport ?? resolveTransportForServerUrl(config.serverUrl);
			if (
				!config.externalSession &&
				transport === 'websocket' &&
				isMixedContentBlocked(config.serverUrl)
			) {
				throw new Error('Mixed-content connection blocked');
			}
			setStatus('connecting');
			const live = await openSession(config, transport);
			if (token !== generation) {
				live.dispose();
				return;
			}
			session = live;
			patch({ active: true, transport, clientId: live.awareness.clientID ?? null });
			deps.onOpen?.(live, config, transport);
			if (token === generation) {
				wireProvider(live, transport);
			}
		} catch {
			fail(token);
		}
	};

	const retry = async (): Promise<void> => {
		if (lastInput) {
			await start(lastInput);
		}
	};

	const disposeTeardown =
		deps.registerTeardown === false
			? () => {}
			: registerCollaborationTeardown({
					leaveOnBeforeUnload: () => !store.getSnapshot().prepared?.externalSession,
					leave: stop,
					rejoin: () => void retry(),
				});

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		start,
		stop,
		retry,
		setStatus,
		getSession: () => session,
		gate,
		destroy() {
			if (destroyed) {
				return;
			}
			disposeTeardown();
			stop();
			destroyed = true;
		},
	};
}
