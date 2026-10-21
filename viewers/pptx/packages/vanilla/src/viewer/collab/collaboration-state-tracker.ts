import type { CollaborationConfig, CollaborationState, ConnectionStatus } from 'ooxml-ui/pptx';
import { buildCollaborationState, createCollaborationStateReporter } from 'ooxml-ui/pptx';

import type { Store, ViewerState } from '../state';

/** What the controller knows about the session; remote people come from the store. */
interface TrackedSession {
	active: boolean;
	config: CollaborationConfig | null;
	status: ConnectionStatus;
	synced: boolean;
	clientId: number | null;
}

export interface CollaborationStateTracker {
	/** Merge what changed and report the shared state when it differs from the last one. */
	patch(change: Partial<TrackedSession>): void;
	/** The current shared state. */
	get(): CollaborationState;
	dispose(): void;
}

/** Derives the shared `CollaborationState` from the controller and the store's presence. */
export function createCollaborationStateTracker(
	store: Store<ViewerState>,
	onChange: ((state: CollaborationState) => void) | undefined,
): CollaborationStateTracker {
	const tracked: TrackedSession = {
		active: false,
		config: null,
		status: 'disconnected',
		synced: false,
		clientId: null,
	};
	const compute = (): CollaborationState =>
		buildCollaborationState({
			active: tracked.active,
			roomId: tracked.config?.roomId,
			status: tracked.status,
			synced: tracked.synced,
			clientId: tracked.clientId,
			userName: tracked.config?.userName,
			userColor: tracked.config?.userColor,
			remoteUsers: store.get().remotePresences,
		});
	const reporter = createCollaborationStateReporter((state) => onChange?.(state));
	const refresh = (): void => reporter.update(compute());
	const unsubscribe = store.subscribe((state, previous) => {
		if (state.remotePresences !== previous.remotePresences) {
			refresh();
		}
	});
	return {
		patch(change) {
			Object.assign(tracked, change);
			refresh();
		},
		get: compute,
		dispose: unsubscribe,
	};
}
