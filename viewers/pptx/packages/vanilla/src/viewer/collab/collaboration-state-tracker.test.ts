import type { CollaborationConfig, CollaborationState } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import type { Store, ViewerState } from '../state';
import { createCollaborationStateTracker } from './collaboration-state-tracker';

const store = {
	get: () => ({ remotePresences: [] }),
	subscribe: () => () => {},
} as unknown as Store<ViewerState>;

const config = (overrides: Partial<CollaborationConfig>): CollaborationConfig =>
	({ roomId: 'room', serverUrl: '', userName: 'Ada', ...overrides }) as CollaborationConfig;

describe('collaboration state tracker write-back', () => {
	it('reports the owner write-back status, idle first, and none for a collaborator', () => {
		const onChange = vi.fn<(state: CollaborationState) => void>();
		const tracker = createCollaborationStateTracker(store, onChange);
		tracker.patch({
			active: true,
			config: config({ role: 'owner', onWriteBack: () => {} }),
			status: 'connected',
			synced: true,
		});
		expect(tracker.get().writeBack).toStrictEqual({ status: 'idle' });
		tracker.patch({ writeBack: { status: 'failed', error: 'offline' } });
		expect(tracker.get().writeBack).toStrictEqual({ status: 'failed', error: 'offline' });
		expect(onChange.mock.calls.at(-1)?.[0].writeBack).toStrictEqual({
			status: 'failed',
			error: 'offline',
		});
		tracker.patch({ config: config({ role: 'collaborator', onWriteBack: () => {} }) });
		expect(tracker.get().writeBack).toBeUndefined();
		tracker.dispose();
	});
});
