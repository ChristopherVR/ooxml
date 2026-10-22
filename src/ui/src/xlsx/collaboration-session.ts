// The collab session `<xlsx-editor>` makes when the host gives a room rather than a session,
// through the shared `openCollaborationSession`, with spreadsheet presence validation.
import type { CollabSession } from 'ooxml-core/collab';
import type { XlsxPresence } from 'ooxml-core/xlsx/collab';
import type { XlsxCollaborationOptions } from './collaboration-types';

/** The two runtime modules sharing needs; loaded only when sharing starts. */
export type CollabRuntime = [
	typeof import('ooxml-core/collab'),
	typeof import('ooxml-core/xlsx/collab'),
];

export function createRoomSession(
	collab: CollabRuntime[0],
	xlsx: CollabRuntime[1],
	options: XlsxCollaborationOptions,
	defaultName: string,
): CollabSession<XlsxPresence> {
	return collab.openCollaborationSession<XlsxPresence>(options, {
		defaultName,
		broadcastPrefix: 'ooxml-xlsx:',
		initialPresence: {},
		sanitizePayload: xlsx.sanitizeXlsxPresence,
	}).session;
}
