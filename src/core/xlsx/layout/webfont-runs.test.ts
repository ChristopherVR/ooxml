import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import type { CellStyle, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { resolveThemeTypeface, workbookFontFamily } from './font-stack';
import { collectWorkbookWebfontRuns } from './webfont-runs';

function styled(workbook: Workbook, font: CellStyle['font']): number {
	const base = workbook.styles[0]!;
	workbook.styles.push({ ...base, font });
	return workbook.styles.length - 1;
}

describe('collectWorkbookWebfontRuns', () => {
	it('lists the Normal font through the theme even in an empty workbook', () => {
		const workbook = createWorkbook();
		workbook.theme.minorFont = 'Aptos Narrow';
		expect(collectWorkbookWebfontRuns(workbook)).toEqual([
			{ family: 'Aptos Narrow', bold: false, italic: false, text: '' },
		]);
	});

	it('reads the fonts of cells with content, rich-text runs and their characters', () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		const arialBold = styled(workbook, { name: 'Arial', bold: true });
		const georgia = styled(workbook, { name: 'Georgia', italic: true });
		putCell(sheet, 0, 0, { value: 'Ab c', styleId: arialBold });
		// A styled but empty cell paints no text.
		putCell(sheet, 1, 0, { value: null, styleId: georgia });
		// Runs merge over the cell font, as the painter does (`cell-view.ts`).
		const plain = styled(workbook, { name: 'Calibri' });
		putCell(sheet, 2, 0, {
			value: 'xy',
			styleId: plain,
			richText: [{ text: 'x' }, { text: 'y', font: { name: 'Courier New' } }],
		});
		putCell(sheet, 3, 0, { value: 12.5 });
		const runs = collectWorkbookWebfontRuns(workbook);
		expect(runs).toContainEqual({ family: 'Arial', bold: true, italic: false, text: 'Abc' });
		expect(runs.some((run) => run.family === 'Georgia')).toBe(false);
		expect(runs).toContainEqual({ family: 'Courier New', bold: false, italic: false, text: 'y' });
		expect(runs).toContainEqual({ family: 'Calibri', bold: false, italic: false, text: 'x12.5' });
	});

	it('reads shape and text box runs, resolving theme typefaces', () => {
		const workbook = createWorkbook();
		workbook.theme.majorFont = 'Cambria';
		workbook.sheets[0]!.drawings.push({
			kind: 'unsupported',
			description: 'shape',
			anchor: { type: 'absolute', from: { col: 0, row: 0, colOff: 0, rowOff: 0 } } as never,
			shape: {
				kind: 'group',
				frame: { x: 0, y: 0, width: 1, height: 1 },
				geometry: 'rect',
				children: [
					{
						kind: 'shape',
						frame: { x: 0, y: 0, width: 1, height: 1 },
						geometry: 'rect',
						text: {
							text: 'Hi',
							paragraphs: [{ runs: [{ text: 'Hi', typeface: '+mj-lt', bold: true }] }],
						},
					},
				],
			},
		});
		expect(collectWorkbookWebfontRuns(workbook)).toContainEqual({
			family: 'Cambria',
			bold: true,
			italic: false,
			text: 'Hi',
		});
	});

	it('stops after the cell budget', () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		const times = styled(workbook, { name: 'Times New Roman' });
		for (let row = 0; row < 10; row++) putCell(sheet, row, 0, { value: 'a' });
		putCell(sheet, 10, 0, { value: 'b', styleId: times });
		const runs = collectWorkbookWebfontRuns(workbook, { maxCells: 10 });
		expect(runs.some((run) => run.family === 'Times New Roman')).toBe(false);
	});
});

describe('workbookFontFamily', () => {
	it('keeps the shared workbook stacks, which list the downloaded clone', () => {
		expect(workbookFontFamily('Calibri')).toBe(
			'"Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
		);
		expect(workbookFontFamily(undefined)).toBe(workbookFontFamily('Calibri'));
		expect(workbookFontFamily('  ')).toBe(workbookFontFamily('Calibri'));
	});

	it('places a verified clone right after a family the shared stack does not know', () => {
		expect(workbookFontFamily('Helvetica')).toBe(
			'"Helvetica", "Arimo", "Calibri", "Carlito", "Segoe UI", sans-serif',
		);
	});

	it('leaves a family without a clone on the shared fallback', () => {
		expect(workbookFontFamily('Fancy Font')).toBe(
			'"Fancy Font", "Calibri", "Carlito", "Segoe UI", sans-serif',
		);
	});
});

describe('resolveThemeTypeface', () => {
	const theme = { majorFont: 'Major', minorFont: 'Minor' };
	it('maps theme references and keeps explicit typefaces', () => {
		expect(resolveThemeTypeface('+mj-lt', theme)).toBe('Major');
		expect(resolveThemeTypeface('+mn-ea', theme)).toBe('Minor');
		expect(resolveThemeTypeface('Arial', theme)).toBe('Arial');
		expect(resolveThemeTypeface(undefined, theme)).toBeUndefined();
	});
});
