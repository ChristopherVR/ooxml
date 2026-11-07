import { describe, expect, it } from 'vitest';
import { collectWebfontFaces } from '../text/webfonts';
import type { VisioDocument, VisioShape, VisioText, VisioTextRun } from './model';
import { collectVisioWebfontRuns } from './webfont-runs';

const run = (
	text: string,
	fontFamily: string,
	extra: Partial<VisioTextRun> = {},
): VisioTextRun => ({
	text,
	fontFamily,
	fontSize: 12 / 72,
	color: '#000000',
	bold: false,
	italic: false,
	underline: false,
	...extra,
});

function text(runs: VisioTextRun[], extra: Partial<VisioText> = {}): VisioText {
	return {
		plainText: runs.map((item) => item.text).join(''),
		runs,
		fontFamily: 'Calibri',
		fontSize: 12 / 72,
		color: '#000000',
		horizontalAlign: 'center',
		verticalAlign: 'middle',
		transform: [1, 0, 0, 1, 0, 0],
		width: 1,
		height: 1,
		margins: { left: 0, right: 0, top: 0, bottom: 0 },
		...extra,
	} as VisioText;
}

function shape(id: string, body: VisioText, extra: Partial<VisioShape> = {}): VisioShape {
	return { id, name: id, hidden: false, children: [], text: body, ...extra } as VisioShape;
}

function drawing(pages: VisioShape[][], masters: VisioShape[][] = []): VisioDocument {
	return {
		format: 'vsdx',
		diagnostics: [],
		pages: pages.map((shapes, index) => ({ id: String(index), name: `Page-${index}`, shapes })),
		masters: masters.map((shapes, index) => ({ id: `m${index}`, name: `M${index}`, shapes })),
	} as unknown as VisioDocument;
}

const byFace = (runs: ReturnType<typeof collectVisioWebfontRuns>) =>
	Object.fromEntries(runs.map((item) => [`${item.family}|${item.bold}|${item.italic}`, item.text]));

describe('collectVisioWebfontRuns', () => {
	it('reads every page, group member, bullet and master preview with its face', () => {
		const grouped = shape('g', text([]), {
			children: [shape('child', text([run('Kid', 'Georgia', { italic: true })]))],
		});
		const bulleted = shape(
			'b',
			text([run('Item', 'Arial', { bold: true })], {
				paragraphs: [
					{
						start: 0,
						end: 4,
						horizontalAlign: 'left',
						indentLeft: 0,
						indentRight: 0,
						indentFirst: 0,
						spaceBefore: 0,
						spaceAfter: 0,
						lineSpacing: { kind: 'multiple', value: 1.2 },
						direction: 'ltr',
						bullet: { text: '•', fontFamily: 'Symbol', fontSize: 0.1, offset: 0 },
					},
				],
			}),
		);
		const runs = collectVisioWebfontRuns(
			drawing(
				[
					[
						shape('a', text([run('Hello ', 'Calibri'), run('Bold', 'Calibri', { bold: true })])),
						grouped,
					],
					[bulleted],
				],
				[[shape('m', text([run('Master', 'Cambria')]))]],
			),
		);
		expect(byFace(runs)).toEqual({
			'Calibri|false|false': 'Helo',
			'Calibri|true|false': 'Bold',
			'Georgia|false|true': 'Kid',
			'Arial|true|false': 'Item',
			'Symbol|false|false': '•',
			'Cambria|false|false': 'Master',
		});
	});

	it('uses the text block defaults when a shape has no runs, and skips hidden shapes', () => {
		const plain = shape('p', {
			...text([]),
			plainText: 'Plain',
			fontFamily: 'Aptos',
			bold: true,
		});
		const hidden = shape('h', text([run('Secret', 'Tinos')]), { hidden: true });
		const runs = collectVisioWebfontRuns(drawing([[plain, hidden]]));
		expect(byFace(runs)).toEqual({ 'Aptos|true|false': 'Plain' });
	});

	it('can leave out the master previews', () => {
		const runs = collectVisioWebfontRuns(
			drawing([[]], [[shape('m', text([run('M', 'Cambria')]))]]),
			{
				includeMasters: false,
			},
		);
		expect(runs).toEqual([]);
	});

	it('bounds the sample per face and the characters scanned', () => {
		const long = Array.from({ length: 600 }, (_, index) => String.fromCodePoint(0x4e00 + index));
		const runs = collectVisioWebfontRuns(
			drawing([[shape('a', text([run(long.join(''), 'Calibri')]))]]),
		);
		expect([...runs[0]!.text]).toHaveLength(256);
		const scanned = collectVisioWebfontRuns(
			drawing([
				[shape('a', text([run('abcdef', 'Calibri')])), shape('b', text([run('xyz', 'Arial')]))],
			]),
			{ maxScannedCharacters: 4 },
		);
		// Scanning stops, but every face the drawing uses is still named.
		expect(byFace(scanned)).toEqual({ 'Calibri|false|false': 'abcd', 'Arial|false|false': '' });
	});

	it('adds the upper-case forms a caps transform draws, merging family spellings', () => {
		const runs = collectVisioWebfontRuns(
			drawing([
				[
					shape('a', text([run('ab', 'Calibri', { textCase: 'all-caps' })])),
					shape('b', text([run('c', ' calibri ')])),
				],
			]),
		);
		expect(byFace(runs)).toEqual({ 'Calibri|false|false': 'abABc' });
	});

	it('feeds the shared face collector, which requests the metric clone', () => {
		const runs = collectVisioWebfontRuns(drawing([[shape('a', text([run('Hi', 'Calibri')]))]]));
		const { faces } = collectWebfontFaces(runs, ['Calibri']);
		expect([...faces.keys()]).toEqual(['Carlito']);
		expect([...faces.get('Carlito')![0]!.chars].join('')).toBe('Hi');
	});
});
