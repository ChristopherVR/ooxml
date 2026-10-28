import type {
	CollabLoadOrigin,
	CollabYSession,
	CollaborationConfig,
	CollaborationTransport,
	ConnectionStatus,
	YDocLike,
	YjsFactories,
} from 'ooxml-ui/pptx';
import {
	bindCollaborationHistory,
	collaborationHistoryReady,
	createCollaborationLivePatcher,
	createCollaborationSessionController,
	createSyncGate,
	DEFAULT_CURSOR_COLOR,
	LOCAL_SYNC_ORIGIN,
	observeYDocSlides,
	readSlidesFromYDoc,
	shouldRoomSlidesReplaceLoad,
} from 'ooxml-ui/pptx';

import type {
	CollaborationController,
	CollaborationControllerDeps,
} from './collaboration-controller-types';
import { createCollaborationEditState } from './collaboration-edit-state';
import type { ExternalSessionSync } from './collaboration-external-sync';
import { createExternalSessionSync } from './collaboration-external-sync';
import { collabModules } from './collaboration-modules';
import type { PresenceController } from './collaboration-presence';
import { createPresenceController } from './collaboration-presence';
import { createCollaborationSession } from './collaboration-session';
import { createSlidesSync } from './collaboration-slides-sync';
import type { SlidesSync } from './collaboration-slides-sync';
import { createCollaborationStateTracker } from './collaboration-state-tracker';
import { createWriteBackScheduler } from './collaboration-writeback';

/**
 * Yjs collaboration with built-in transports or a borrowed host session. The
 * room lifecycle (owned sessions, transport, connection status, first-write
 * gate, pagehide teardown) is the shared `createCollaborationSessionController`;
 * granular slide sync and presence live in their focused modules.
 *
 * Undo and Redo go through the room's own history (this editor's operations only, so a
 * participant's later edit survives) while local edits are being written into the room.
 */

export function createCollaborationController(
	deps: CollaborationControllerDeps,
): CollaborationController {
	const { store } = deps;

	let status: ConnectionStatus = 'disconnected';
	let active = false;
	let externalSync: ExternalSessionSync | null = null;
	let currentYDoc: YDocLike | null = null;
	let yFactories: YjsFactories | null = null;
	let presence: PresenceController | null = null;
	const editState = createCollaborationEditState(deps);
	let lastConfig: CollaborationConfig | null = null;
	let unobserveSlides: (() => void) | null = null;
	let unsubscribeStore: (() => void) | null = null;
	let loadApplying = false;
	let releaseHistory: (() => void) | null = null;
	const livePatcher = createCollaborationLivePatcher();
	const stateTracker = createCollaborationStateTracker(store, deps.onStateChange);

	/** Bind (or release) the room history as the connection, sync and role change. */
	function updateHistoryBinding(): void {
		const history = deps.getHistory?.();
		const ready =
			history !== undefined &&
			collaborationHistoryReady({
				hasDoc: currentYDoc !== null,
				external: Boolean(lastConfig?.externalSession),
				connected: status === 'connected',
				synced: syncGate.isOpen(),
				viewer: lastConfig?.role === 'viewer',
			});
		if (ready && !releaseHistory && currentYDoc) {
			releaseHistory = bindCollaborationHistory(history, currentYDoc, collabModules.loadYjs);
		} else if (!ready && releaseHistory) {
			releaseHistory();
			releaseHistory = null;
		}
	}

	function setStatus(next: ConnectionStatus): void {
		if (next === status) {
			return;
		}
		status = next;
		stateTracker.patch({ status: next });
		updateHistoryBinding();
		deps.onStatusChange?.(next);
	}

	const writeBack = createWriteBackScheduler({
		getYDoc: () => currentYDoc,
		getHandler: deps.getHandler,
		getSaveOptions: deps.getSaveOptions,
		onWriteBackStatus: () => stateTracker.patch({ writeBack: writeBack.getState() }),
	});

	const slidesSync: SlidesSync = createSlidesSync(store, (config) => {
		if (!config.externalSession || syncGate.isOpen()) {
			writeBack.schedule(config);
		}
	});

	function flushLocal(): void {
		slidesSync.flushLocalSlides(currentYDoc, yFactories, lastConfig, editState.isReadOnly());
	}
	// Prevent a bootstrap deck from preceding initial sync. Built-in transports
	// retain their grace timer; external sessions only use host readiness.
	const baseGate = createSyncGate(() => {
		stateTracker.patch({ synced: true });
		flushLocal();
		updateHistoryBinding();
	});
	const syncGate = {
		...baseGate,
		reset: () => {
			baseGate.reset();
			stateTracker.patch({ synced: false });
		},
	};

	/** The session is live: attach presence, the slide mapping and the readiness owner. */
	function attach(
		created: CollabYSession,
		config: CollaborationConfig,
		_transport: CollaborationTransport,
	): void {
		currentYDoc = created.doc;
		yFactories = created.factories;
		presence = createPresenceController(
			store,
			created.awareness,
			{
				userName: config.userName,
				userColor: config.userColor ?? DEFAULT_CURSOR_COLOR,
				userAvatar: config.userAvatar,
				role: config.role,
			},
			() => ({ width: store.get().canvasSize.width, height: store.get().canvasSize.height }),
		);

		// Observe remote slide changes, skipping our own reconcile transactions.
		if (!config.externalSession) {
			unobserveSlides = observeYDocSlides(currentYDoc, (_events, transaction) => {
				if (transaction?.origin === LOCAL_SYNC_ORIGIN || slidesSync.isApplyingRemote()) {
					return;
				}
				slidesSync.applyRemoteSlides(currentYDoc, config);
			});
		}

		// Broadcast local slide edits granularly (diff by id, one transaction).
		// Suppressed until the sync gate opens (the gate flushes on open) and
		// while the load pipeline is committing a parsed deck (adoption in
		// notifyContentLoaded decides whether that deck may be published).
		unsubscribeStore = store.subscribe((state, previous) => {
			if (
				state.slides !== previous.slides &&
				!state.deckStyleHeld &&
				!previous.deckStyleHeld &&
				!loadApplying &&
				syncGate.isOpen()
			) {
				flushLocal();
			}
		});

		active = true;
		stateTracker.patch({ active: true, config, clientId: created.awareness.clientID ?? null });
		updateHistoryBinding();
		if (config.externalSession) {
			externalSync = createExternalSessionSync(config.externalSession, config, {
				factories: created.factories,
				livePatcher,
				gate: syncGate,
				setStatus: (next) => session.setStatus(next),
				setReadOnly: editState.setReadOnly,
				flushLocal,
				resetBaseline: slidesSync.reset,
				cancelWriteBack: writeBack.cancel,
				adoptSlides: (slides) => slidesSync.adoptSlides(slides, config),
			});
		} else {
			livePatcher.configure(
				editState.isReadOnly() ? null : currentYDoc,
				editState.isReadOnly() ? null : yFactories,
			);
		}
	}

	/** The session is ending: detach everything `attach` set up. */
	function detach(): void {
		releaseHistory?.();
		releaseHistory = null;
		externalSync?.dispose();
		externalSync = null;
		loadApplying = false;
		writeBack.cancel();
		slidesSync.reset();
		unobserveSlides?.();
		unobserveSlides = null;
		unsubscribeStore?.();
		unsubscribeStore = null;
		presence?.destroy();
		presence = null;
		currentYDoc = null;
		yFactories = null;
		livePatcher.configure(null, null);
		active = false;
		stateTracker.patch({ active: false, config: null, clientId: null });
		editState.setReadOnly(false);
	}

	const session = createCollaborationSessionController({
		openSession: createCollaborationSession,
		gate: syncGate,
		onBegin: (config) => {
			lastConfig = config;
			editState.setReadOnly(Boolean(config.externalSession) || config.role === 'viewer');
		},
		onOpen: attach,
		onClose: detach,
	});
	session.subscribe(() => setStatus(session.getSnapshot().status));

	// Loads suppress publishing until adoption decides whether the room's deck
	// beats a late bootstrap. Explicit user-opened files instead replace the room.
	function beginContentLoad(_origin: CollabLoadOrigin): void {
		loadApplying = true;
	}

	function notifyContentLoaded(origin: CollabLoadOrigin): void {
		const suppressed = loadApplying;
		loadApplying = false;
		if (!active || !currentYDoc || !lastConfig) {
			return;
		}
		if (externalSync) {
			externalSync.contentLoaded(origin);
			return;
		}
		if (
			shouldRoomSlidesReplaceLoad(origin, readSlidesFromYDoc(currentYDoc).length) &&
			slidesSync.applyRemoteSlides(currentYDoc, lastConfig)
		) {
			return;
		}
		if (suppressed && syncGate.isOpen()) {
			flushLocal();
		}
	}

	return {
		start: (input) => session.start(input),
		stop: () => session.stop(),
		isActive: () => active,
		isReadOnly: editState.isReadOnly,
		getStatus: () => status,
		getState: stateTracker.get,
		setCursor: (x, y, activeSlideIndex) => presence?.setCursor(x, y, activeSlideIndex),
		setSelection: (selectedElementId, activeSlideIndex) =>
			presence?.setSelection(selectedElementId, activeSlideIndex),
		setActiveSlide: (index) => presence?.setActiveSlide(index),
		followUser: (clientId) => presence?.followUser(clientId ?? null),
		getConfig: () => lastConfig,
		beginContentLoad,
		notifyContentLoaded,
		livePatcher,
		destroy: () => {
			session.destroy();
			stateTracker.dispose();
		},
	};
}
