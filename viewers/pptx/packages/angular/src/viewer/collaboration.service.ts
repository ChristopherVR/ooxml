/**
 * CollaborationService: Angular real-time collaboration (Yjs) service.
 *
 * Owns viewer state and built-in transports, or borrows a host-owned session:
 *  - The room lifecycle (owned sessions, y-websocket or serverless y-webrtc,
 *    mixed-content and connect-timeout errors, retry, pagehide teardown) is the
 *    shared createCollaborationSessionController. External resources are never
 *    destroyed.
 *  - Local edits reconcile changed slides/elements/fields in a transaction
 *    tagged LOCAL_SYNC_ORIGIN; observers skip their own local writes.
 *  - Elected-writer write-back (role owner) uses the shared WriteBackScheduler.
 *
 * Provide at the component level: `@Component({ providers: [CollaborationService] })`.
 */

import { DestroyRef, Injectable, computed, inject, signal } from '@angular/core';
import type {
	CollabLoadOrigin,
	CollabYSession,
	CollaborationHistoryTarget,
	CollaborationState,
	CollaborationConfig,
	CollaborationLivePatcher,
	CollaborationRole,
	ConnectionStatus,
} from 'ooxml-ui/pptx';
import {
	bindCollaborationHistory,
	buildCollaborationState,
	collaborationHistoryReady,
	createCollaborationLivePatcher,
	createCollaborationSessionController,
	writeBackReport,
} from 'ooxml-ui/pptx';
import type { PptxHandlerSaveOptions, PptxSlide } from 'pptx-viewer-core';

import { activateExternalSession } from './collaboration-external-session';
import { CollaborationPresenceState } from './collaboration-presence-state';
import { collabModules, openCollaborationSession } from './collaboration-providers';
import { activateSession } from './collaboration-session-setup';
import type { ActiveSession, ConnectOptions } from './collaboration-session-setup';
import { SlideSyncEngine } from './collaboration-slide-sync';
import { WriteBackScheduler } from './collaboration-writeback';
import { controllerSignal } from './controller-signal';
import type { TemplateElementsBySlideId } from './template-mode';

// Re-exported so existing importers of `ConnectOptions` from this module keep
// resolving after the interface moved to collaboration-session-setup.ts.
export type { ConnectOptions } from './collaboration-session-setup';

/** Sentinel canvas bound used until the host reports real dimensions. */
const DEFAULT_CANVAS_BOUND = 100_000;

@Injectable()
export class CollaborationService {
	// Reactive state
	readonly status = computed<ConnectionStatus>(() => this.sessionState().status);
	readonly connected = computed<boolean>(() => this.status() === 'connected');
	readonly active = computed<boolean>(() => this.sessionState().active);
	/** Role of the local user in the active session, or undefined when idle. */
	readonly activeRole = signal<CollaborationRole | undefined>(undefined);
	/** Active viewer role or host readiness can temporarily block editing. */
	readonly readOnly = signal(false);
	private readonly presenceState = new CollaborationPresenceState();
	readonly presence = this.presenceState.presence;
	readonly cursors = this.presenceState.cursors;
	readonly connectedCount = computed<number>(
		() => this.presence().length + (this.active() ? 1 : 0),
	);

	readonly followedClientId = this.presenceState.followedClientId;
	readonly followedSlideIndex = this.presenceState.followedSlideIndex;
	readonly broadcasterSlideIndex = this.presenceState.broadcasterSlideIndex;

	/**
	 * Interim ("live preview") Y.Doc write channel: publishes in-flight inline
	 * editor text that has not reached the slides state yet, so peers see typing
	 * as it happens instead of on commit. Dormant outside a session.
	 */
	readonly livePatcher: CollaborationLivePatcher = createCollaborationLivePatcher();

	// One handle owns viewer wiring, but external transport resources stay borrowed.
	private session: ActiveSession | null = null;
	private readonly writeBackState = signal<CollaborationState['writeBack']>(undefined);
	private readonly writeBack = new WriteBackScheduler(() =>
		this.writeBackState.set(this.writeBack.getState()),
	);
	/** Granular local<->doc slide sync (gate + echo dedupe + broadcast/adopt). */
	private readonly slideSync = new SlideSyncEngine();
	/**
	 * The room lifecycle (owned sessions, transport, connection status, the
	 * first-write gate, pagehide teardown, reentrancy), shared by every binding.
	 */
	private readonly sessionCtl = createCollaborationSessionController({
		openSession: (config, transport) => openCollaborationSession(config, transport),
		gate: this.slideSync.gate,
		onBegin: (config) => this.begin(config),
		onOpen: (live, config) => this.attach(live, config),
		onClose: () => this.detach(),
	});
	private readonly sessionState = controllerSignal(this.sessionCtl);
	/** The first exchange with the room finished (the first-write gate is open). */
	readonly synced = this.slideSync.synced;
	/** The local awareness client id of the active session, or null. */
	readonly clientId = signal<number | null>(null);
	private readonly activeConfig = signal<CollaborationConfig | null>(null);
	/** The shared `CollaborationState` a host reads (`collaborationChange`). */
	readonly state = computed<CollaborationState>(() => {
		const config = this.activeConfig();
		return buildCollaborationState({
			active: this.active(),
			roomId: config?.roomId,
			status: this.status(),
			synced: this.synced(),
			clientId: this.clientId(),
			userName: config?.userName,
			userColor: config?.userColor,
			remoteUsers: this.presence(),
			writeBack: writeBackReport(config, this.writeBackState()),
		});
	});

	private onRemoteSlides: ((slides: PptxSlide[]) => void) | null = null;
	private canvasWidth = DEFAULT_CANVAS_BOUND;
	private canvasHeight = DEFAULT_CANVAS_BOUND;
	private getSourceBytes: (() => Uint8Array | null) | null = null;
	private getTemplateElements: (() => TemplateElementsBySlideId) | null = null;
	private getSaveOptions: (() => PptxHandlerSaveOptions) | null = null;
	private currentConfig: CollaborationConfig | null = null;
	private history: CollaborationHistoryTarget | null = null;
	private readonly liveDoc = signal<unknown>(null);
	private releaseHistory: (() => void) | null = null;
	private lastOptions: ConnectOptions = {};

	private readonly refreshPresence = (): void => {
		this.presenceState.refresh(this.session, this.canvasWidth, this.canvasHeight);
	};

	constructor() {
		// The session controller also leaves the room from `pagehide`.
		inject(DestroyRef).onDestroy(() => this.sessionCtl.destroy());
		this.slideSync.onGateChange = () => this.updateHistoryBinding();
		this.sessionCtl.subscribe(() => this.updateHistoryBinding());
	}

	/** Undo and Redo reach only this user's edits while they are written into the room. */
	private updateHistoryBinding(): void {
		const history = this.history;
		const doc = this.liveDoc();
		const config = this.activeConfig();
		const ready =
			history !== null &&
			collaborationHistoryReady({
				hasDoc: doc !== null,
				external: Boolean(config?.externalSession),
				connected: this.status() === 'connected',
				synced: this.synced(),
				viewer: this.activeRole() === 'viewer',
			});
		if (ready && !this.releaseHistory && doc) {
			this.releaseHistory = bindCollaborationHistory(history, doc, collabModules.loadYjs);
		} else if (!ready && this.releaseHistory) {
			this.releaseHistory();
			this.releaseHistory = null;
		}
	}
	async connect(requested: CollaborationConfig, options: ConnectOptions = {}): Promise<void> {
		this.lastOptions = options;
		this.history = options.history ?? null;
		await this.sessionCtl.start(requested);
	}

	/** Reconnect using the configuration from the most recent {@link connect}. */
	async retry(): Promise<void> {
		await this.sessionCtl.retry();
	}

	disconnect(): void {
		this.sessionCtl.stop();
	}

	/** A session is starting with `config` (again once a room prepared it). */
	private begin(config: CollaborationConfig): void {
		const options = this.lastOptions;
		this.onRemoteSlides = options.onRemoteSlides ?? null;
		this.canvasWidth = options.canvasWidth ?? DEFAULT_CANVAS_BOUND;
		this.canvasHeight = options.canvasHeight ?? DEFAULT_CANVAS_BOUND;
		this.getSourceBytes = options.getSourceBytes ?? null;
		this.getTemplateElements = options.getTemplateElements ?? null;
		this.getSaveOptions = options.getSaveOptions ?? null;
		// The requested config stays the session's identity; the role comes from the prepared one.
		const requested = this.sessionCtl.getSnapshot().config ?? config;
		this.currentConfig = requested;
		this.activeConfig.set(requested);
		this.activeRole.set(config.role);
		this.readOnly.set(Boolean(config.externalSession) || config.role === 'viewer');
	}

	/** The session is live: bind the slide sync, presence and (borrowed) readiness. */
	private attach(live: CollabYSession, config: CollaborationConfig): void {
		const deps = {
			slideSync: this.slideSync,
			livePatcher: this.livePatcher,
			onRemoteSlides: this.onRemoteSlides,
			refreshPresence: this.refreshPresence,
			scheduleWriteBack: () => this.scheduleWriteBack(),
			cancelWriteBack: () => this.writeBack.cancel(),
			setReadOnly: (readOnly: boolean) => this.readOnly.set(readOnly),
			setStatus: (status: ConnectionStatus) => this.sessionCtl.setStatus(status),
		};
		const session = config.externalSession
			? activateExternalSession(config.externalSession, config, live, deps)
			: activateSession(live, config, deps);
		this.session = session;
		this.liveDoc.set(live.doc);
		this.updateHistoryBinding();
		this.clientId.set(session.selfId);
		this.refreshPresence();
		// Readiness can seed/adopt synchronously before this.session exists.
		// Now an owner's startup snapshot has its live document available.
		if (session.readiness) {
			this.scheduleWriteBack();
		}
	}

	/** The session is ending (also on every restart): detach before it is disposed. */
	private detach(): void {
		this.liveDoc.set(null);
		this.updateHistoryBinding();
		this.slideSync.reset();
		this.writeBack.cancel();
		this.writeBackState.set(undefined);
		this.session?.dispose();
		this.session = null;
		this.livePatcher.configure(null, null);
		this.onRemoteSlides = null;
		this.currentConfig = null;
		this.activeConfig.set(null);
		this.clientId.set(null);
		this.activeRole.set(undefined);
		this.readOnly.set(false);
		this.presenceState.reset();
	}

	/**
	 * Record the current local deck as the sync baseline so the first (unchanged)
	 * broadcast after connecting is suppressed. Call right after {@link connect}
	 * for a joiner whose local deck is a placeholder awaiting remote sync.
	 */
	seedBaseline(slides: readonly PptxSlide[]): void {
		this.slideSync.seedBaseline(slides);
	}

	/**
	 * Re-adopt the shared document's slides after a local content load committed
	 * a parsed deck to viewer state (see {@link SlideSyncEngine.adoptDocAfterLoad}).
	 * Returns true when the room's slides were adopted over the loaded deck.
	 */
	adoptDocSlidesAfterLoad(origin: CollabLoadOrigin = 'user'): boolean {
		if (this.session?.readiness) {
			return this.session.readiness.handleLoad(origin);
		}
		return this.connected() ? this.slideSync.adoptDocAfterLoad(origin) : false;
	}

	/**
	 * Broadcast the local slide set to peers, reconciling only what changed into
	 * the pptx:slides Y.Array. Empty/unchanged decks are skipped; while the gate
	 * is shut the deck is held pending until the initial sync confirms.
	 */
	broadcastSlides(slides: readonly PptxSlide[]): void {
		this.slideSync.broadcast(slides);
	}

	setCursor(x: number, y: number, activeSlideIndex?: number): void {
		this.session?.localPresence.setCursor(x, y, activeSlideIndex);
	}

	setSelection(selectedElementId: string | undefined, activeSlideIndex?: number): void {
		this.session?.localPresence.setSelection(selectedElementId, activeSlideIndex);
	}

	/** Publish the local active-slide index (drives follow-along). */
	setActiveSlide(index: number): void {
		this.session?.localPresence.setActiveSlide(index);
	}

	/** Follow the given peer's active slide, or null to stop following. */
	followUser(clientId: number | null): void {
		this.followedClientId.set(clientId);
	}

	private scheduleWriteBack(): void {
		if (this.currentConfig?.externalSession && !this.slideSync.gate.isOpen()) {
			return;
		}
		this.writeBack.schedule(
			this.currentConfig,
			this.session?.ydoc ?? null,
			this.getSourceBytes,
			this.getTemplateElements,
			this.getSaveOptions,
		);
	}
}
