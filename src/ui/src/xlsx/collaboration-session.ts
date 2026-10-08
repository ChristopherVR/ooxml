// The collab session `<xlsx-editor>` makes when the host gives a room rather than a session: a
// provider of the host's, a y-websocket compatible server, or a BroadcastChannel between the
// windows of this browser, with spreadsheet presence validation and pagehide teardown.
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
	const roomId = options.roomId ?? '';
	if (!collab.isValidRoomId(roomId))
		throw new Error('Use 1 to 128 letters, digits, hyphens or underscores for the room name.');
	let provider = options.provider;
	if (!provider && options.serverUrl)
		provider = collab.transportProvider({
			transport: collab.createWebSocketTransport({
				url: collab.roomUrl(options.serverUrl, roomId),
			}),
		});
	if (!provider) {
		if (!collab.canBroadcast())
			throw new Error('This browser cannot share between windows; pass a server URL.');
		provider = collab.transportProvider({
			transport: collab.createBroadcastTransport({ roomId, prefix: 'ooxml-xlsx:' }),
		});
	}
	const user = options.user ?? {};
	return collab.createCollabSession<XlsxPresence>({
		roomId,
		provider,
		user: {
			name: user.name || defaultName,
			...(user.color ? { color: user.color } : {}),
			...(user.role ? { role: user.role } : {}),
		},
		initialPresence: {},
		sanitizePayload: xlsx.sanitizeXlsxPresence,
		teardown: true,
	});
}
