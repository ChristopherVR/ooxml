import { describe, expect, it } from 'vitest';
import { visioBackgroundFieldShapes } from './background-fields';
import type { VisioPage, VisioShape, VisioTextRun } from './model';
import { demoDocument } from './ui/demo-document';

const run = (text: string, bold = false): VisioTextRun => ({
	text,
	fontFamily: 'Arial',
	fontSize: 0.15,
	color: '#000000',
	bold,
	italic: false,
	underline: false,
});

/** A background page whose footer reads "Page: <PAGENAME()> of <PAGENUMBER()>". */
function background(): { page: VisioPage; footer: VisioShape } {
	const page = structuredClone(demoDocument.pages[0]!);
	page.id = 'bg';
	page.name = 'Background-1';
	page.isBackground = true;
	const footer = page.shapes[0]!;
	const plainText = 'Page: Background-1 #1';
	footer.text = {
		...footer.text!,
		plainText,
		runs: [run('Page: '), run('Background-1', true), run(' #1')],
		paragraphs: [
			{
				start: 0,
				end: plainText.length,
				horizontalAlign: 'center',
				indentLeft: 0,
				indentRight: 0,
				indentFirst: 0,
				spaceBefore: 0,
				spaceAfter: 0,
				lineSpacing: { kind: 'multiple', value: 1 },
				direction: 'ltr',
			},
		],
		fields: [
			{ start: 6, end: 18, cached: 'Background-1', formula: 'PAGENAME()' },
			{ start: 20, end: 21, cached: '1', formula: 'PAGENUMBER()' },
		],
	};
	page.shapes = [footer, page.shapes[2]!];
	return { page, footer };
}

describe('page fields on a background page', () => {
	it('show the foreground page that uses the background (Visio PAGENAME semantics)', () => {
		const { page, footer } = background();
		const [shape, other] = visioBackgroundFieldShapes(page, { name: 'Orders', number: 3 });
		expect(shape).not.toBe(footer);
		expect(shape!.text!.plainText).toBe('Page: Orders #3');
		expect(shape!.text!.runs.map((entry) => entry.text)).toEqual(['Page: ', 'Orders', ' #3']);
		expect(shape!.text!.runs[1]!.bold).toBe(true);
		expect(shape!.text!.fields!.map(({ start, end }) => [start, end])).toEqual([
			[6, 12],
			[14, 15],
		]);
		expect(shape!.text!.paragraphs![0]!.end).toBe('Page: Orders #3'.length);
		// Shapes without page fields are kept as they are, and the source page is unchanged.
		expect(other).toBe(page.shapes[1]);
		expect(footer.text!.plainText).toBe('Page: Background-1 #1');
	});

	it('keeps the shapes when nothing changes', () => {
		const { page } = background();
		expect(visioBackgroundFieldShapes(page, { name: 'Background-1', number: 1 })).toBe(page.shapes);
	});

	it('follows an upper-case picture and leaves other formulas alone', () => {
		const { page, footer } = background();
		footer.text!.plainText = 'Page: BACKGROUND-1 #1';
		footer.text!.runs = [run('Page: BACKGROUND-1 #1')];
		footer.text!.fields![1]!.formula = 'PAGENUMBER()+1';
		const [shape] = visioBackgroundFieldShapes(page, { name: 'Orders', number: 3 });
		expect(shape!.text!.plainText).toBe('Page: ORDERS #1');
		expect(shape!.text!.runs.map((entry) => entry.text)).toEqual(['Page: ORDERS #1']);
	});

	it.each([
		['PAGENAME()', true],
		['=pagename( 0 )', true],
		['= PAGENAME(0)\t', true],
		['PAGENAME(\t\t)', true],
		['PAGENAME(00)', false],
		['PAGENAME(1)', false],
		['PAGENAME(0', false],
	])('reads %j as a page name field: %s', (formula, isPageName) => {
		const { page, footer } = background();
		footer.text!.fields![0]!.formula = formula;
		const [shape] = visioBackgroundFieldShapes(page, { name: 'Orders', number: 1 });
		expect(shape!.text!.plainText).toBe(isPageName ? 'Page: Orders #1' : 'Page: Background-1 #1');
	});

	it('reads a long page name formula that does not close in linear time', () => {
		const { page, footer } = background();
		footer.text!.fields![0]!.formula = `PAGENAME(${'\t'.repeat(100_000)}x`;
		const started = performance.now();
		visioBackgroundFieldShapes(page, { name: 'Orders', number: 1 });
		expect(performance.now() - started).toBeLessThan(500);
	});
});
