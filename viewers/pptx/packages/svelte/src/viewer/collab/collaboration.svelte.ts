/**
 * Real-time collaboration for the Svelte viewer (Yjs: y-websocket or serverless
 * y-webrtc). The room lifecycle is the shared `createCollaborationSessionController`
 * (read through `controllerState`); the remote-slide observer lives in
 * `collaboration-remote-sync.ts` and presence in `collaboration-presence.svelte.ts`.
 * KNOWN LIMITATION: collaborative-undo semantics are undefined in shared - local
 * undo is kept as-is and may fight a concurrent remote edit (matching the others).
 */
import type {
	CollaborationConfig,
	CollaborationSessionController,
	CollaborationSessionState,
	CollaborationState,
	CollabLoadOrigin,
	CollabYSession,
	ConnectionStatus,
	RemoteCursor,
	SanitizedPresence,
} from 'ooxml-ui/pptx';
import {
	buildCollaborationState,
	createCollaborationSessionController,
	writeBackReport,
	DEFAULT_CURSOR_COLOR,
} from 'ooxml-ui/pptx';

import { controllerState } from '../state/controller-state';
import type { ControllerState } from '../state/controller-state';
import type { CollaborationDeps } from './collaboration-deps';
import { CollaborationDocument } from './collaboration-document.svelte';
import { registerCollaborationEffects } from './collaboration-effects.svelte';
import { CollaborationPresence } from './collaboration-presence.svelte';
import type { CollabSessionFactory } from './collaboration-session';
import {
	createDefaultSession,
	createExternalSession,
	toSharedSession,
} from './collaboration-session';

/**
 * The collaboration controller. Construct it once during component setup: it
 * registers the effects (auto start/stop from the config, granular publish of
 * local edits) itself, so no further wiring is needed in the SFC.
 */
export class CollaborationController {
	/** Interim Y.Doc channel for in-flight inline text (dormant when stopped). */
	get livePatcher() {
		return this.#document.livePatcher;
	}

	readonly #deps: CollaborationDeps;
	readonly #makeSession: CollabSessionFactory;
	readonly #session: CollaborationSessionController;
	readonly #state: ControllerState<CollaborationSessionState>;
	#lastStarted: CollaborationConfig | null = null;
	#startedByEffect = false;
	readonly #document: CollaborationDocument;
	readonly #presence: CollaborationPresence;

	constructor(deps: CollaborationDeps) {
		this.#deps = deps;
		this.#document = new CollaborationDocument(deps, (status) => {
			this.status = status;
		});
		this.#makeSession = deps.createSession ?? createDefaultSession;
		this.#presence = new CollaborationPresence(() => ({
			width: this.#deps.getCanvasWidth?.(),
			height: this.#deps.getCanvasHeight?.(),
		}));
		this.#session = createCollaborationSessionController({
			openSession: async (config, transport) =>
				config.externalSession
					? createExternalSession(config.externalSession)
					: toSharedSession(await this.#makeSession(transport, config)),
			gate: this.#document.gate,
			onBegin: (config) => this.#document.begin(config),
			onOpen: (live, config) => this.#attach(live, config),
			onClose: () => this.#detach(),
		});
		this.#state = controllerState(this.#session);

		registerCollaborationEffects({
			getConfig: () => this.#deps.getConfig(),
			getSlides: () => this.#deps.getSlides(),
			syncConfig: (config) => this.#syncConfig(config),
			isPublishable: () => this.active && this.#document.gate.isOpen(),
			flushLocalSlides: (slides) => this.#document.flush(slides),
			destroy: () => this.#session.destroy(),
		});
	}

	/** Live connection status (reactive). */
	get status(): ConnectionStatus {
		return this.#state.current.status;
	}
	/** A borrowed host session's readiness reports its status here. */
	set status(next: ConnectionStatus) {
		this.#session.setStatus(next);
	}
	/** Whether a session is live (reactive). */
	get active(): boolean {
		return this.#state.current.active;
	}
	/** The requested viewer role forbids editing, including while attachment is pending. */
	get readOnly(): boolean {
		return this.#document.readOnly;
	}
	/** Remote cursors on the current slide (reactive). */
	get cursors(): RemoteCursor[] {
		return this.#presence.cursors;
	}
	/** Remote collaborators in the session (reactive). */
	get remotePresences(): SanitizedPresence[] {
		return this.#presence.remotePresences;
	}
	/** Followed peer's client id, or null when free (reactive). */
	get followedClientId(): number | null {
		return this.#presence.followedClientId;
	}
	/** The config the active session was started with (null when stopped); the
	 * Share dialog's active view reads the local user's name/colour from this. */
	get activeCollaboration(): CollaborationConfig | null {
		return this.#state.current.config;
	}
	/** The shared `CollaborationState` a host reads (`oncollaborationchange`), reactive. */
	get state(): CollaborationState {
		const config = this.activeCollaboration;
		return buildCollaborationState({
			active: this.active,
			roomId: config?.roomId,
			status: this.status,
			synced: this.#document.synced,
			clientId: this.#presence.selfClientId,
			userName: config?.userName,
			userColor: config ? (config.userColor ?? DEFAULT_CURSOR_COLOR) : undefined,
			remoteUsers: this.remotePresences,
			writeBack: writeBackReport(config, this.#document.writeBack),
		});
	}
	/** Total connected participants (self + remote peers), reactive. */
	get connectedCount(): number {
		return this.remotePresences.length + (this.active ? 1 : 0);
	}

	/** Publish a cursor move (slide-space px); no-op when no session is active. */
	setCursor(x: number, y: number, activeSlideIndex?: number): void {
		this.#presence.setCursor(x, y, activeSlideIndex);
	}
	/** Publish the local selection; no-op when no session is active. */
	setSelection(selectedElementId: string | undefined, activeSlideIndex?: number): void {
		this.#presence.setSelection(selectedElementId, activeSlideIndex);
	}
	/** Publish the local active-slide index (drives peer follow-along). */
	setActiveSlide(index: number): void {
		this.#presence.setActiveSlide(index);
	}
	/** Follow the given peer's active slide, or `null` to stop following. */
	followUser(clientId: number | null): void {
		this.#presence.followUser(clientId);
	}

	/**
	 * Re-adopt the shared doc's slides after a local content load committed a
	 * parsed deck to viewer state (see `adoptDocSlidesAfterLoad`). The load
	 * pipeline calls this synchronously right after it applies, i.e. before the
	 * publish effect can flush the freshly loaded slides into the doc, so a
	 * late joiner's bootstrap deck never clobbers the room's synced content.
	 */
	adoptDocAfterLoad(origin: CollabLoadOrigin = 'user'): void {
		if (this.active) {
			this.#document.adoptAfterLoad(origin);
		}
	}

	#syncConfig(config: CollaborationConfig | undefined): void {
		if (config && config !== this.#lastStarted) {
			this.#lastStarted = config;
			this.#startedByEffect = true;
			void this.#session.start(config);
		} else if (!config && this.#startedByEffect) {
			// Only auto-stop a session THIS effect started; a direct `start()`
			// call (e.g. from a dialog) always clears the flag below, so it
			// is immune to this branch on the effect's next run.
			this.#lastStarted = null;
			this.#startedByEffect = false;
			this.stop();
		}
	}

	/** Start (or restart) a session with the given config (dialog-driven). */
	async start(config: CollaborationConfig): Promise<void> {
		// Set synchronously, before any `await` below, so a same-tick effect
		// flush (see `#syncConfig`) sees this config as already current and
		// does not redundantly start a second, concurrent session.
		this.#lastStarted = config;
		this.#startedByEffect = false;
		await this.#session.start(config);
	}

	/** The session is live: attach the document and presence. */
	#attach(live: CollabYSession, config: CollaborationConfig): void {
		this.#document.attach(live.doc, live.factories);
		this.#presence.start(live.awareness, {
			userName: config.userName,
			userColor: config.userColor ?? DEFAULT_CURSOR_COLOR,
			userAvatar: config.userAvatar,
			role: config.role,
		});
		this.#deps.onStart?.(config);
	}

	/** The session is ending (also on every restart): detach before it is disposed. */
	#detach(): void {
		this.#document.stop();
		this.#presence.stop();
		// The controller is still being built when its first stop() runs.
		if (this.#session?.getSnapshot().active) {
			this.#deps.onStop?.();
		}
	}

	stop(): void {
		this.#session.stop();
	}
}
