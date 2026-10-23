import type {
	CollaborationConfig,
	ExternalCollaborationSession,
	ExternalCollaborationReadiness,
	CollabYSession,
} from 'ooxml-ui/pptx';
import { DEFAULT_CURSOR_COLOR, observeExternalCollaborationReadiness } from 'ooxml-ui/pptx';

import { LocalPresencePublisher } from './collaboration-local-presence';
import type { ActiveSession, ActivateSessionDeps } from './collaboration-session-setup';

/**
 * Attach viewer-owned observers to a borrowed host session without acquiring
 * ownership of host resources: the shared session controller already holds the
 * presence lease (`live.awareness`) and releases it on stop.
 */
export function activateExternalSession(
	external: ExternalCollaborationSession,
	config: CollaborationConfig,
	live: CollabYSession,
	deps: ActivateSessionDeps,
): ActiveSession {
	const factories = live.factories;
	const awareness = live.awareness as ActiveSession['awareness'] & {
		off: (event: string, cb: () => void) => void;
	};
	deps.slideSync.bind({
		ydoc: external.doc,
		factories,
		onRemoteSlides: deps.onRemoteSlides,
		scheduleWriteBack: deps.scheduleWriteBack,
		readOnly: config.role === 'viewer',
		external: true,
	});
	const localPresence = new LocalPresencePublisher(awareness, {
		userName: config.userName,
		userColor: config.userColor ?? DEFAULT_CURSOR_COLOR,
		userAvatar: config.userAvatar,
		role: config.role,
	});
	localPresence.publish();
	awareness.on('change', deps.refreshPresence);
	awareness.on('update', deps.refreshPresence);
	let readiness: ExternalCollaborationReadiness | undefined;
	const dispose = (): void => {
		readiness?.();
		awareness.off('change', deps.refreshPresence);
		awareness.off('update', deps.refreshPresence);
	};
	try {
		readiness = observeExternalCollaborationReadiness(external, {
			gate: deps.slideSync.gate,
			livePatcher: deps.livePatcher,
			factories,
			role: config.role,
			sessionIntent: config.sessionIntent,
			onStatus: deps.setStatus,
			onReadOnlyChange: deps.setReadOnly,
			adoptSlides: (slides) => deps.slideSync.adoptSlides(slides),
			onReady: ({ seedEmptyRoom }) => deps.slideSync.publishCurrent(seedEmptyRoom),
			onSuspend: () => {
				deps.slideSync.cancelPending();
				deps.cancelWriteBack();
			},
		});
	} catch (error) {
		dispose();
		throw error;
	}
	return {
		ydoc: external.doc,
		awareness,
		selfId: awareness.clientID ?? -1,
		localPresence,
		readiness,
		dispose,
	};
}
