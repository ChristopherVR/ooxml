// The binding on real collab sessions over the in-memory transport: gate, seeding, presence.
import { afterEach, describe, expect, it } from 'vitest';
import { createBroadcastTransport } from '../../collab/broadcast-transport';
import { createMemoryHub } from '../../collab/memory-transport';
import { type CollabSession, createCollabSession } from '../../collab/session';
import { transportProvider } from '../../collab/transport-provider';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { createWorkbook } from '../workbook';
import { bindWorkbookSession } from './bind';
import { type XlsxPresence, sanitizeXlsxPresence } from './presence';

const open: CollabSession<XlsxPresence>[] = [];
afterEach(() => {
	for (const session of open.splice(0)) session.destroy();
});

type Hub = ReturnType<typeof createMemoryHub> | { broadcast: string };

function join(hub: Hub, name: string, sheets: string[], syncGraceMs = 0) {
	const transport =
		'broadcast' in hub
			? createBroadcastTransport({ roomId: hub.broadcast })
			: hub.createTransport('book');
	const session = createCollabSession<XlsxPresence>({
		roomId: 'book',
		provider: transportProvider({ transport }),
		user: { name },
		initialPresence: {},
		sanitizePayload: sanitizeXlsxPresence,
		heartbeatMs: 0,
		syncGraceMs,
	});
	open.push(session);
	const edit = createEditSession(createWorkbook({ sheets }));
	return { session, edit, binding: bindWorkbookSession(session, edit) };
}

describe('bindWorkbookSession on collab sessions', () => {
	it('seeds once, adopts the room on join and shares edits and selections', async () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada', ['Budget']);
		await new Promise((resolve) => setTimeout(resolve, 5));
		a.edit.setCellInput(0, 0, 0, '42');
		const b = join(hub, 'Bob', ['Ignored']);
		await new Promise((resolve) => setTimeout(resolve, 5));
		expect(b.edit.workbook.sheets.map((s) => s.name)).toEqual(['Budget']);
		expect(getCell(b.edit.workbook.sheets[0]!, 0, 0)?.value).toBe(42);
		b.edit.setCellInput(0, 1, 0, '=A1+1');
		expect(getCell(a.edit.workbook.sheets[0]!, 1, 0)).toMatchObject({ formula: 'A1+1', value: 43 });
		b.binding.setSelection(0, { start: { row: 1, col: 0 }, end: { row: 1, col: 0 } });
		await new Promise((resolve) => setTimeout(resolve, 80));
		expect(a.binding.remoteSelections()).toMatchObject([{ userName: 'Bob', sheet: 0 }]);
	});

	it('a guest joining before the first peer seeded adopts that peer workbook', async () => {
		// Ada is alone and still inside her grace period (nothing written yet) when Bob arrives: the
		// handshake syncs Bob to an empty room and Ada's awareness arrives before any workbook. The
		// BroadcastChannel delivers Bob's opening sync before his hello, as in the browser.
		const hub = { broadcast: `race-${Math.random().toString(36).slice(2)}` };
		const a = join(hub, 'Ada', ['Sales', 'Budget'], 150);
		a.edit.setCellInput(0, 0, 0, 'from Ada');
		await new Promise((resolve) => setTimeout(resolve, 20));
		const b = join(hub, 'Bob', ['Sheet1'], 150);
		await new Promise((resolve) => setTimeout(resolve, 400));
		for (const side of [a, b]) {
			expect(side.edit.workbook.sheets.map((s) => s.name)).toEqual(['Sales', 'Budget']);
			expect(getCell(side.edit.workbook.sheets[0]!, 0, 0)?.value).toBe('from Ada');
		}
	});

	it('peers that seed at the same moment converge to one workbook', async () => {
		const hub = { broadcast: `same-${Math.random().toString(36).slice(2)}` };
		const a = join(hub, 'Ada', ['Sales'], 60);
		const b = join(hub, 'Bob', ['Notes'], 60);
		await new Promise((resolve) => setTimeout(resolve, 300));
		const names = (side: typeof a) => side.edit.workbook.sheets.map((s) => s.name);
		expect(names(a)).toEqual(names(b));
		expect(names(a).length).toBeGreaterThan(0);
	});
});
