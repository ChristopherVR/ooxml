import { createBuildingBlockShell, resolveCollaborationConfig } from 'ooxml-ui/pptx';

import type { BuildingBlocksViewer } from './blocks-viewer';
import type { ViewerBuildingBlocksOptions } from './types';

/**
 * The shared building-block shell (`shellState`, `collaborationState`) over the mounted vanilla
 * viewer's store and collaboration session; idle until a canvas mounts.
 */
export function createVanillaShell(
	viewer: () => BuildingBlocksViewer | null,
	options: () => ViewerBuildingBlocksOptions,
) {
	return createBuildingBlockShell(() => {
		const mounted = viewer();
		if (!mounted) {
			return null;
		}
		const state = mounted.store.get();
		const host = options();
		const collaborationState = mounted.getCollaborationState();
		return {
			shell: {
				authorizedCanEdit: host.canEdit === true,
				configured:
					collaborationState.active || Boolean(resolveCollaborationConfig(host.collaboration)),
				readOnly: mounted.collaborationReadOnly(),
				sourcePending: Boolean(host.content) && state.loading,
				sourceError: Boolean(state.error),
				status: mounted.getCollaborationStatus(),
				remoteUsers: state.remotePresences,
			},
			collaboration: collaborationState,
		};
	});
}
