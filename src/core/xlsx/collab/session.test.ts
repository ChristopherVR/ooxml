// The binding on real collab sessions over the in-memory transport: gate, seeding, presence.
import { afterEach, describe, expect, it } from 'vitest';
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

function join(hub: ReturnType<typeof createMemoryHub>, name: string, sheets: string[]) {
	const session = createCollabSession<XlsxPresence>({
		roomId: 'book',
		provider: transportProvider({ transport: hub.createTransport('book') }),
		user: { name },
		initialPresence: {},
		sanitizePayload: sanitizeXlsxPresence,
		heartbeatMs: 0,
		syncGraceMs: 0,
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
});
