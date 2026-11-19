/**
 * share-helpers.ts: Share dialog helpers for the Angular viewer.
 *
 * The shared subset (ShareFormFields / ShareDefaults / seedShareFields /
 * canStartShare) is re-exported from `pptx-viewer-shared` (`render/share-form`);
 * links come from the shared `buildCollaborationShareUrl`. One helper stays local
 * because it diverges from the shared builders: `buildCollaborationConfig`
 * validates with `canStartShare` (non-blank room) and emits a config WITHOUT
 * `role`/`sessionIntent`, unlike the shared `buildShareConfig` (which is
 * `buildCreateCollaborationConfig`).
 *
 * No `any`; all regexes use the `/u` flag; no `String.prototype.replaceAll`,
 * no regex named-capture-groups (ng-packagr lib-target constraints).
 */

import type { CollaborationConfig, ShareDefaults, ShareFormFields } from 'ooxml-ui/pptx';
import {
	canStartShare,
	createShareSecret,
	resolveTransportForServerUrl,
	seedShareFields,
} from 'ooxml-ui/pptx';

export type { ShareDefaults, ShareFormFields };
export { canStartShare, seedShareFields };

/**
 * Assemble a {@link CollaborationConfig} from the (trimmed) form fields, or
 * `null` when the form is incomplete. A blank server URL yields
 * `transport: 'webrtc'` with a fresh session secret; otherwise the default
 * websocket transport is used.
 */
export function buildCollaborationConfig(fields: ShareFormFields): CollaborationConfig | null {
	if (!canStartShare(fields)) {
		return null;
	}
	const serverUrl = fields.serverUrl.trim();
	const transport = resolveTransportForServerUrl(serverUrl);
	return {
		roomId: fields.roomId.trim(),
		userName: fields.userName.trim(),
		serverUrl,
		transport,
		...(transport === 'webrtc' ? { secret: createShareSecret() } : {}),
	};
}
