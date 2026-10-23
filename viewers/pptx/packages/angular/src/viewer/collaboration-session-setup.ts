/**
 * collaboration-session-setup.ts: attach the Angular `CollaborationService`'s
 * document and presence wiring to a live collaboration session.
 *
 * The session itself (document, transport, connection status, first-write
 * gate) comes from the shared `createCollaborationSessionController`;
 * `activateSession` binds the slide-sync engine and presence to it and returns
 * the {@link ActiveSession} the service holds as a single atomic handle.
 */

import type {
	CollabYSession,
	CollaborationConfig,
	CollaborationLivePatcher,
	ConnectionStatus,
	ExternalCollaborationReadiness,
	YDocLike,
} from 'ooxml-ui/pptx';
import { observeYDocSlides } from 'ooxml-ui/pptx';
import type { PptxHandlerSaveOptions, PptxSlide } from 'pptx-viewer-core';

import { DEFAULT_CURSOR_COLOR } from './collaboration-helpers';
import { LocalPresencePublisher } from './collaboration-local-presence';
import type { AwarenessLike } from './collaboration-providers';
import type { SlideSyncEngine } from './collaboration-slide-sync';
import type { TemplateElementsBySlideId } from './template-mode';

/** Options a host passes to `connect`, held for reconnection. */
export interface ConnectOptions {
	onRemoteSlides?: (slides: PptxSlide[]) => void;
	canvasWidth?: number;
	canvasHeight?: number;
	getSourceBytes?: () => Uint8Array | null;
	/**
	 * Returns the editor's separated template (master/layout) elements keyed by
	 * slide id, so the elected-writer write-back can merge them back into the
	 * broadcast (template-free) slides before serializing. Without this, template
	 * edits would be dropped from the persisted deck.
	 */
	getTemplateElements?: () => TemplateElementsBySlideId;
	/**
	 * Session-level save options (view properties, table styles, tags, deck
	 * properties, ...), built the same way as the Save/Export path
	 * (`buildDeckSaveOptions`). Without this the elected-writer write-back
	 * dropped every session-level edit outside `slides`.
	 */
	getSaveOptions?: () => PptxHandlerSaveOptions;
}

/** The transport objects + wiring handles of one live session, owned together. */
export interface ActiveSession {
	ydoc: YDocLike;
	awareness: AwarenessLike;
	selfId: number;
	localPresence: LocalPresencePublisher;
	dispose: () => void;
	readiness?: ExternalCollaborationReadiness;
}

/** Everything `activateSession` reads from / calls back into the service. */
export interface ActivateSessionDeps {
	slideSync: SlideSyncEngine;
	livePatcher: CollaborationLivePatcher;
	onRemoteSlides: ((slides: PptxSlide[]) => void) | null;
	/** Recompute the presence signal from awareness (bound in the service). */
	refreshPresence: () => void;
	/** Schedule an owner-role write-back after a doc mutation. */
	scheduleWriteBack: () => void;
	cancelWriteBack: () => void;
	setReadOnly: (readOnly: boolean) => void;
	/** A borrowed session's readiness reports the connection status here. */
	setStatus: (status: ConnectionStatus) => void;
}

/**
 * Wire an owned session into the service: configure the live-patch channel,
 * bind the slide-sync engine, publish local presence and subscribe to awareness
 * and remote-slide changes. The connection status, the first-write gate and the
 * transport teardown belong to the shared session controller.
 */
export function activateSession(
	live: CollabYSession,
	config: CollaborationConfig,
	deps: ActivateSessionDeps,
): ActiveSession {
	const viewer = config.role === 'viewer';
	deps.livePatcher.configure(viewer ? null : live.doc, viewer ? null : live.factories);
	deps.setReadOnly(viewer);
	deps.slideSync.bind({
		ydoc: live.doc,
		factories: live.factories,
		readOnly: viewer,
		onRemoteSlides: deps.onRemoteSlides,
		scheduleWriteBack: deps.scheduleWriteBack,
	});

	const awareness = live.awareness as AwarenessLike;
	const localPresence = new LocalPresencePublisher(awareness, {
		userName: config.userName,
		userColor: config.userColor ?? DEFAULT_CURSOR_COLOR,
		userAvatar: config.userAvatar,
		role: config.role,
	});
	localPresence.publish();
	awareness.on('change', deps.refreshPresence);
	awareness.on('update', deps.refreshPresence);

	const unobserve = observeYDocSlides(live.doc, (_events, transaction) =>
		deps.slideSync.onRemoteChange(transaction),
	);

	return {
		ydoc: live.doc,
		awareness,
		selfId: awareness.clientID ?? -1,
		localPresence,
		dispose() {
			unobserve();
			awareness.off?.('change', deps.refreshPresence);
			awareness.off?.('update', deps.refreshPresence);
		},
	};
}
