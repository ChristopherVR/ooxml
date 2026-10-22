import type { CollaborationState } from 'ooxml-core/collab';
import {
	buildCollaborationState,
	createCollaborationStateReporter,
	writeBackReport,
} from 'ooxml-ui/pptx';
import { useEffect, useMemo, useRef } from 'react';

import type { CollaborationContextValue } from './types';

/**
 * Reports the shared `CollaborationState` of the active session (or the off state when
 * `collab` is null) to `onChange`, once per real change. Call it where the collaboration
 * context is available.
 */
export function useCollaborationStateReport(
	collab: CollaborationContextValue | null,
	onChange: ((state: CollaborationState) => void) | undefined,
): void {
	const callback = useRef(onChange);
	callback.current = onChange;
	const reporter = useMemo(
		() => createCollaborationStateReporter((state) => callback.current?.(state)),
		[],
	);
	const status = collab?.status;
	const synced = collab?.synced;
	const clientId = collab?.clientId;
	const remoteUsers = collab?.remoteUsers;
	const config = collab?.config;
	const writeBack = collab?.writeBack;
	useEffect(() => {
		reporter.update(
			buildCollaborationState({
				active: Boolean(config),
				roomId: config?.roomId,
				status: status ?? 'disconnected',
				synced: synced ?? false,
				clientId,
				userName: config?.userName,
				userColor: config?.userColor,
				remoteUsers: remoteUsers ?? [],
				writeBack: writeBackReport(config, writeBack),
			}),
		);
	}, [reporter, status, synced, clientId, remoteUsers, config, writeBack]);
}
