import type {
	CollabLoadOrigin,
	CollaborationConfig,
	CollaborationLivePatcher,
	ConnectionStatus,
	ExternalCollaborationSession,
	SyncGate,
	YjsFactories,
} from 'ooxml-ui/pptx';
import { observeExternalCollaborationReadiness } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

interface ExternalSyncDeps {
	factories: YjsFactories;
	livePatcher: CollaborationLivePatcher;
	gate: SyncGate;
	adoptSlides(slides: PptxSlide[]): void;
	/** The adopted room slides replaced the open deck (a join): a presentation load. */
	onRoomReplacedDeck(): void;
	flushLocal(): void;
	resetBaseline(): void;
	cancelWriteBack(): void;
	setStatus(status: ConnectionStatus): void;
	setReadOnly(readOnly: boolean): void;
}

export interface ExternalSessionSync {
	contentLoaded(origin: CollabLoadOrigin): void;
	dispose(): void;
}

/** Host readiness replaces transport grace timers, including the initial join barrier. */
export function createExternalSessionSync(
	session: ExternalCollaborationSession,
	config: CollaborationConfig,
	deps: ExternalSyncDeps,
): ExternalSessionSync {
	const readiness = observeExternalCollaborationReadiness(session, {
		gate: deps.gate,
		factories: deps.factories,
		livePatcher: deps.livePatcher,
		role: config.role,
		sessionIntent: config.sessionIntent,
		adoptSlides: deps.adoptSlides,
		onRoomReplacedDeck: deps.onRoomReplacedDeck,
		onStatus: deps.setStatus,
		onReadOnlyChange: deps.setReadOnly,
		onSuspend: deps.cancelWriteBack,
		onReady: ({ seedEmptyRoom }) => {
			if (seedEmptyRoom) {
				deps.resetBaseline();
			}
			deps.flushLocal();
		},
	});
	return {
		contentLoaded(origin) {
			if (!readiness.handleLoad(origin) && readiness.canWrite()) {
				deps.flushLocal();
			}
		},
		dispose: readiness,
	};
}
