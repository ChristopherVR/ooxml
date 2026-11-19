import { createHash } from 'node:crypto';

/**
 * The demo URL that joins the serverless (webrtc) room `roomId`, the way a Share dialog
 * invitation does: the session secret travels in the `#key=` fragment. The secret is derived
 * from the room id so every page of one test meets in the same session.
 */
export function webrtcRoomPath(roomId: string, query: string): string {
	const secret = createHash('sha256').update(roomId).digest().subarray(0, 16).toString('base64url');
	return `/?${query}room=${encodeURIComponent(roomId)}&transport=webrtc#key=${secret}`;
}
