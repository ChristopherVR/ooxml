import { describe, expect, it } from 'vitest';
import { createEditSession, createWorkbook, getCell } from '../../index';
import type { CellRange } from '../../index';
import type { EditorContext } from '../context';
import { dataCommands } from './data';
import { DATA_TOOLS_STRINGS } from './i18n/data-tools';

function makeContext(range: CellRange = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }) {
	const wb = createWorkbook();
	const session = createEditSession(wb);
	const toasts: { message: string; kind?: string }[] = [];
	const opened: string[] = [];
	let currentRange = range;
	const ctx = {
		workbook: () => wb,
		session: () => session,
		activeSheet: () => 0,
		readOnly: () => false,
		selection: {
			get: () => ({
				sheet: 0,
				active: currentRange.start,
				anchor: currentRange.start,
				ranges: [currentRange],
			}),
		},
		dialogs: {
			open: async (name: string) => {
				opened.push(name);
				return undefined;
			},
		},
		t: (text: string) => text,
		toast: (message: string, kind?: string) => toasts.push({ message, ...(kind ? { kind } : {}) }),
	} as unknown as EditorContext;
	return {
		wb,
		session,
		ctx,
		toasts,
		opened,
		setRange: (r: CellRange) => (currentRange = r),
	};
}

describe('Data tab commands (Flash Fill, Subtotal, Consolidate, Advanced Filter)', () => {
	it('runs Flash Fill directly and fills inferred values', () => {
		const { wb, session, ctx, toasts } = makeContext({
			start: { row: 0, col: 1 },
			end: { row: 0, col: 1 },
		});
		session.setRangeValues(0, { row: 0, col: 0 }, [
			['Ada Lovelace', 'Ada'],
			['Alan Turing', null],
			['Grace Hopper', null],
		]);
		const cmd = dataCommands().find((c) => c.id === 'data.flash-fill')!;
		expect(cmd.shortcut).toBe('Ctrl+E');
		expect(cmd.enabled?.(ctx)).toBe(true);
		cmd.run(ctx);
		expect(getCell(wb.sheets[0]!, 1, 1)?.value).toBe('Alan');
		expect(getCell(wb.sheets[0]!, 2, 1)?.value).toBe('Grace');
		expect(toasts).toEqual([]);
	});

	it('shows an info toast when Flash Fill finds no pattern', () => {
		const { ctx, toasts } = makeContext({
			start: { row: 0, col: 1 },
			end: { row: 0, col: 1 },
		});
		const cmd = dataCommands().find((c) => c.id === 'data.flash-fill')!;
		cmd.run(ctx);
		expect(toasts).toEqual([{ message: 'No pattern found for Flash Fill.', kind: 'info' }]);
	});

	it('opens Subtotal, Consolidate and Advanced Filter dialogs', () => {
		const { ctx, opened } = makeContext();
		const byId = new Map(dataCommands().map((c) => [c.id, c]));
		byId.get('data.subtotal')!.run(ctx);
		byId.get('data.consolidate')!.run(ctx);
		byId.get('data.advanced-filter')!.run(ctx);
		expect(opened).toEqual(['subtotal', 'consolidate', 'advanced-filter']);
	});

	it('provides all 4 locale translations without em-dashes in DATA_TOOLS_STRINGS', () => {
		for (const [key, texts] of Object.entries(DATA_TOOLS_STRINGS)) {
			expect(texts.length, key).toBe(4);
			for (const text of texts) {
				expect(text.trim().length, key).toBeGreaterThan(0);
				expect(text, key).not.toContain('\u2014');
			}
		}
	});
});
