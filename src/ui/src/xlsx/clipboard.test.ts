import { afterEach, describe, expect, it } from 'vitest';
import { getCell, parseRange } from 'ooxml-core/xlsx';
import { createGridClipboard } from './clipboard.js';
import { createTestContext } from './grid/test-context.js';

const range = (ref: string) => {
	const result = parseRange(ref);
	if (!result) throw new Error(ref);
	return result;
};
afterEach(() => document.body.replaceChildren());
const setup = () => {
	const ctx = createTestContext();
	const session = ctx.session()!;
	const marquees: unknown[] = [];
	const clipboard = createGridClipboard({
		ctx,
		sink: document.createElement('textarea'),
		isEditing: () => false,
		setMarquee: (next) => marquees.push(next),
		selectRange: (next) => ctx.selection.set({ ranges: [next] }),
	});
	return { ctx, session, marquees, clipboard, sheet: session.workbook.sheets[0]! };
};

describe('grid clipboard selection', () => {
	it('pastes a copied value across the selected rectangle', () => {
		const { ctx, session, clipboard, sheet } = setup();
		session.setCellValue(0, 0, 0, 'Monday');
		const payload = clipboard.copy(false)!;
		ctx.selection.set({ ranges: [range('C3:D4')] });
		expect(clipboard.paste({ text: payload.tsv })).toBe(true);
		expect(getCell(sheet, 3, 3)?.value).toBe('Monday');
		expect(ctx.selection.get().ranges).toEqual([range('C3:D4')]);
	});

	it('ends cut mode after the core consumes the move flag', () => {
		const { ctx, session, clipboard, sheet, marquees } = setup();
		session.setCellValue(0, 0, 0, 'move');
		const payload = clipboard.copy(true)!;
		ctx.selection.set({ ranges: [range('C3')] });
		expect(clipboard.paste({ text: payload.tsv })).toBe(true);
		expect(getCell(sheet, 0, 0)).toBeUndefined();
		expect(getCell(sheet, 2, 2)?.value).toBe('move');
		expect(marquees.at(-1)).toBeUndefined();
		expect(clipboard.clearMarquee()).toBe(false);
		expect(clipboard.internal()?.cut).toBe(false);
	});

	it('reports incompatible sizes and leaves the workbook unchanged', () => {
		const { ctx, session, clipboard, sheet } = setup();
		session.setCellValue(0, 0, 0, 1);
		ctx.selection.set({ ranges: [range('A1:B2')] });
		const payload = clipboard.copy(false)!;
		ctx.selection.set({ ranges: [range('C3:E5')] });
		expect(clipboard.paste({ text: payload.tsv })).toBe(false);
		expect(ctx.toasts).toContain('The copy and paste areas are not compatible sizes.');
		expect(getCell(sheet, 2, 2)).toBeUndefined();
	});
});
