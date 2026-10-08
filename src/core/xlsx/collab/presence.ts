// Spreadsheet presence: where a collaborator is, as the shared sheet key (stable across renames
// and moves) and an A1 range. Published through the `collab` session's awareness payload.
import type { RemotePresence } from '../../collab/presence';
import { type CellRange, formatRange, parseRange } from '../address';

/** The awareness payload of a spreadsheet collaborator. Both fields are absent until they select. */
export interface XlsxPresence {
	/** Shared key of the sheet (not its name or index). */
	sheet?: string;
	/** The selection as an A1 range (`B2` or `B2:D9`). */
	range?: string;
	/**
	 * When this peer's binding joined (ms since the epoch). A peer finding the room empty leaves the
	 * seeding to a writer that joined before it, so a guest never replaces that writer's workbook.
	 */
	joined?: number;
}

/** A collaborator's selection resolved against the local workbook. */
export interface RemoteSelection {
	clientId: number;
	userName: string;
	userColor: string;
	/** Index of the sheet in the local workbook. */
	sheet: number;
	range: CellRange;
}

const SHEET_KEY = /^[a-z0-9]{1,16}-[a-z0-9]{1,16}$/;

/**
 * Validates a peer's payload for `createCollabSession({ sanitizePayload })`: unknown or malformed
 * fields are dropped, but the peer stays visible.
 */
export function sanitizeXlsxPresence(raw: Record<string, unknown>): XlsxPresence {
	const out: XlsxPresence = {};
	const { sheet, range, joined } = raw;
	if (typeof joined === 'number' && Number.isFinite(joined) && joined > 0) out.joined = joined;
	if (typeof sheet !== 'string' || !SHEET_KEY.test(sheet)) return out;
	if (typeof range !== 'string' || range.length > 32) return out;
	const parsed = parseRange(range);
	if (!parsed) return out;
	out.sheet = sheet;
	out.range = formatRange(parsed);
	return out;
}

/** Resolves peers' payloads to local sheet indexes; peers on unknown sheets are left out. */
export function resolveSelections(
	peers: readonly RemotePresence<XlsxPresence>[],
	sheetIndex: (key: string) => number | undefined,
): RemoteSelection[] {
	const out: RemoteSelection[] = [];
	for (const peer of peers) {
		if (!peer.sheet || !peer.range) continue;
		const sheet = sheetIndex(peer.sheet);
		const range = parseRange(peer.range);
		if (sheet === undefined || !range) continue;
		out.push({
			clientId: peer.clientId,
			userName: peer.userName,
			userColor: peer.userColor,
			sheet,
			range,
		});
	}
	return out;
}
