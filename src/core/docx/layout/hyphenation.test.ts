import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, twips } from '../index';
import { SOFT_HYPHEN } from '../../text/hyphenation';
import { adaptDocumentModel } from './adapter';
import type { LayoutParagraph, LayoutRun } from './input';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';
import { layoutParagraph } from './paragraph-layout';
import { at } from './test-helpers';
import { tokenPosition } from './text-breaks';

// Monospace measurer: 10px per character, 20px lines.
const measurer: TextMeasurer = {
	widthOf: (text) => text.length * 10,
	lineHeightOf: () => 20,
};
const noop = () => {};
const lineText = (fragments: { text: string }[]) => fragments.map((f) => f.text).join('');
const allLines = (p: LayoutParagraph, widthPx: number) =>
	layoutParagraph(p, widthPx, measurer, noop).lines.map((line) => lineText(line.fragments));

function paragraph(runs: LayoutRun[], overrides: Partial<LayoutParagraph> = {}): LayoutParagraph {
	return { kind: 'paragraph', id: 'p1', runs, ...overrides };
}

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const CONTENT_TYPES =
	'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';

async function buildDocx(
	bodyXml: string,
	settingsXml = '',
	stylesXml?: string,
): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', CONTENT_TYPES);
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="3000" w:h="6000"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`,
	);
	if (settingsXml) zip.file('word/settings.xml', `<w:settings ${NS}>${settingsXml}</w:settings>`);
	if (stylesXml) zip.file('word/styles.xml', `<w:styles ${NS}>${stylesXml}</w:styles>`);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('automatic hyphenation in layoutParagraph', () => {
	it('leaves layout unchanged when autoHyphenation is off', () => {
		const input = paragraph([{ text: 'See formatting' }]);
		expect(allLines(input, 100)).toEqual(['See', 'formatting']);
	});

	it('splits overflowing words with a visible hyphen and preserves source offsets', () => {
		const text = 'See formatting';
		const input = paragraph([{ text }], { hyphenation: { zoneTwips: twips(0) } });
		const result = layoutParagraph(input, 100, measurer, noop);
		expect(result.lines.map((line) => lineText(line.fragments))).toEqual(['See for-', 'matting']);
		// "See " (0..4) + "for" (4..7) => line 0 covers 0..7; line 1 covers "matting" (7..14).
		expect(at(result.lines, 0)).toMatchObject({ sourceStart: 0, sourceEnd: 7 });
		expect(at(result.lines, 1)).toMatchObject({ sourceStart: 7, sourceEnd: 14 });
		// Resuming from line 0's nextToken/nextTokenOffset yields the exact remainder.
		const resumed = layoutParagraph(
			input,
			100,
			measurer,
			noop,
			undefined,
			tokenPosition(at(result.lines, 0).nextToken!, at(result.lines, 0).nextTokenOffset),
		);
		expect(resumed.lines.map((line) => lineText(line.fragments))).toEqual(['matting']);
		expect(at(resumed.lines, 0)).toMatchObject({ sourceStart: 7, sourceEnd: 14 });
	});

	it('respects hyphenationZone: skips hyphenation when remaining gap is within the zone', () => {
		// Line width = 100px. "Seven" = 50px, so remaining gap before "formatting" is 50px.
		// With zone = 60px (900 twips), gap (50px) <= zone (60px) => do NOT hyphenate.
		expect(
			allLines(
				paragraph([{ text: 'Seven formatting' }], { hyphenation: { zoneTwips: twips(900) } }),
				100,
			),
		).toEqual(['Seven', 'formatting']);
		// With zone = 40px (600 twips), gap (50px) > zone (40px) => hyphenate!
		expect(
			allLines(
				paragraph([{ text: 'Seven formatting' }], { hyphenation: { zoneTwips: twips(600) } }),
				100,
			),
		).toEqual(['Seven for-', 'matting']);
	});

	it('uses Word default hyphenationZone (360 twips = 24px) when zoneTwips is omitted', () => {
		// Line width = 100px. "12345678" = 80px => gap = 20px <= 24px => do NOT hyphenate.
		expect(
			allLines(paragraph([{ text: '12345678 formatting' }], { hyphenation: {} }), 100),
		).toEqual(['12345678', 'formatting']);
		// Line width = 100px. "12345" = 50px => gap = 50px > 24px => hyphenate.
		expect(allLines(paragraph([{ text: '12345 formatting' }], { hyphenation: {} }), 100)).toEqual([
			'12345 for-',
			'matting',
		]);
	});

	it('respects consecutiveHyphenLimit and resets the count after a non-hyphenated line', () => {
		// Each word is 10 chars ("formatting"); line width 80px forces overflow on every line.
		const text = 'aa formatting formatting formatting formatting';
		const unlimited = allLines(
			paragraph([{ text }], { hyphenation: { zoneTwips: twips(0), consecutiveLimit: 0 } }),
			80,
		);
		// With no limit, multiple consecutive lines end with '-'.
		const unlimitedHyphenCount = unlimited.filter((l) => l.endsWith('-')).length;
		expect(unlimitedHyphenCount).toBeGreaterThanOrEqual(3);

		// With consecutiveLimit = 1, no two consecutive lines may end with '-'.
		const limit1 = allLines(
			paragraph([{ text }], { hyphenation: { zoneTwips: twips(0), consecutiveLimit: 1 } }),
			80,
		);
		expect(limit1.some((l) => l.endsWith('-'))).toBe(true);
		for (let i = 1; i < limit1.length; i++) {
			expect(at(limit1, i - 1).endsWith('-') && at(limit1, i).endsWith('-')).toBe(false);
		}
	});

	it('respects doNotHyphenateCaps for ALL-CAPS words but still hyphenates Title-Case words', () => {
		expect(
			allLines(
				paragraph([{ text: 'See FORMATTING' }], {
					hyphenation: { zoneTwips: twips(0), doNotHyphenateCaps: true },
				}),
				100,
			),
		).toEqual(['See', 'FORMATTING']);
		expect(
			allLines(
				paragraph([{ text: 'See FORMATTING' }], {
					hyphenation: { zoneTwips: twips(0), doNotHyphenateCaps: false },
				}),
				100,
			),
		).toEqual(['See FOR-', 'MATTING']);
		expect(
			allLines(
				paragraph([{ text: 'See Formatting' }], {
					hyphenation: { zoneTwips: twips(0), doNotHyphenateCaps: true },
				}),
				100,
			),
		).toEqual(['See For-', 'matting']);
	});

	it('hyphenates an oversized single word at a syllable boundary instead of an arbitrary grapheme', () => {
		// "formatting" (100px) in an 80px column: auto-hyphenation splits as "for-mat-" / "ting"
		// rather than arbitrary "formatti" / "ng".
		const lines = allLines(
			paragraph([{ text: 'formatting' }], { hyphenation: { zoneTwips: twips(0) } }),
			80,
		);
		expect(lines).toEqual(['format-', 'ting']);
	});
});

describe('explicit soft hyphens (U+00AD) in layoutParagraph', () => {
	it('hides unbroken soft hyphens with zero width and breaks with a visible hyphen on overflow', () => {
		const text = `co${SOFT_HYPHEN}operate`;
		// Fits on a 90px line ("cooperate" = 9 chars = 90px, soft hyphen takes 0px).
		expect(allLines(paragraph([{ text }]), 90)).toEqual(['cooperate']);
		// Overflows an 80px line: breaks at soft hyphen EVEN WITH autoHyphenation OFF.
		const result = layoutParagraph(paragraph([{ text }]), 80, measurer, noop);
		expect(result.lines.map((line) => lineText(line.fragments))).toEqual(['co-', 'operate']);
		// "co\u00AD" is 3 chars (0..3), "operate" is 7 chars (3..10).
		expect(at(result.lines, 0)).toMatchObject({ sourceStart: 0, sourceEnd: 3 });
		expect(at(result.lines, 1)).toMatchObject({ sourceStart: 3, sourceEnd: 10 });
	});

	it('picks the rightmost fitting soft hyphen and ignores hyphenationZone / doNotHyphenateCaps', () => {
		const text = `IN${SOFT_HYPHEN}TER${SOFT_HYPHEN}NATIONAL`;
		// 100px line: "IN-TER-"? No, "INTER-" (60px) fits, while "INTERNATIONAL" (130px) overflows.
		const lines = allLines(
			paragraph([{ text }], {
				hyphenation: { zoneTwips: twips(3000), doNotHyphenateCaps: true },
			}),
			100,
		);
		expect(lines).toEqual(['INTER-', 'NATIONAL']);
	});
});

describe('DOCX settings & suppressAutoHyphens end-to-end layout', () => {
	it('reads settings.xml hyphenation options and honours paragraph/style suppressAutoHyphens', async () => {
		const styles =
			'<w:style w:type="paragraph" w:styleId="NoHyphen"><w:pPr><w:suppressAutoHyphens/></w:pPr></w:style>';
		const body = [
			'<w:p><w:r><w:t>See formatting</w:t></w:r></w:p>',
			'<w:p><w:pPr><w:suppressAutoHyphens/></w:pPr><w:r><w:t>See formatting</w:t></w:r></w:p>',
			'<w:p><w:pPr><w:pStyle w:val="NoHyphen"/></w:pPr><w:r><w:t>See formatting</w:t></w:r></w:p>',
			'<w:p><w:pPr><w:pStyle w:val="NoHyphen"/><w:suppressAutoHyphens w:val="0"/></w:pPr><w:r><w:t>See formatting</w:t></w:r></w:p>',
			'<w:p><w:r><w:t>See FORMATTING</w:t></w:r></w:p>',
		].join('');
		const settings =
			'<w:autoHyphenation/><w:hyphenationZone w:val="0"/><w:consecutiveHyphenLimit w:val="2"/><w:doNotHyphenateCaps/>';

		const { model } = await loadDocx(await buildDocx(body, settings, styles));
		expect(model.autoHyphenation).toBe(true);
		expect(model.hyphenationZoneTwips).toBe(twips(0));
		expect(model.consecutiveHyphenLimit).toBe(2);
		expect(model.doNotHyphenateCaps).toBe(true);

		const layoutDoc = adaptDocumentModel(model);
		const paragraphs = at(layoutDoc.sections, 0).blocks.filter(
			(b): b is LayoutParagraph => b.kind === 'paragraph',
		);
		expect(at(paragraphs, 0).hyphenation).toEqual({
			zoneTwips: twips(0),
			consecutiveLimit: 2,
			doNotHyphenateCaps: true,
		});
		expect(at(paragraphs, 1).hyphenation).toBeUndefined();
		expect(at(paragraphs, 2).hyphenation).toBeUndefined();
		expect(at(paragraphs, 3).hyphenation).toBeDefined();
		expect(at(paragraphs, 4).hyphenation).toBeDefined();

		// Verify rendered lines in a 100px-wide layout:
		expect(allLines(at(paragraphs, 0), 100)).toEqual(['See for-', 'matting']);
		expect(allLines(at(paragraphs, 1), 100)).toEqual(['See', 'formatting']);
		expect(allLines(at(paragraphs, 2), 100)).toEqual(['See', 'formatting']);
		expect(allLines(at(paragraphs, 3), 100)).toEqual(['See for-', 'matting']);
		expect(allLines(at(paragraphs, 4), 100)).toEqual(['See', 'FORMATTING']);

		// Full document layout runs cleanly with w:softHyphen too.
		const softDoc = await loadDocx(
			await buildDocx('<w:p><w:r><w:t>co</w:t><w:softHyphen/><w:t>operate</w:t></w:r></w:p>', ''),
		);
		const laid = layoutDocumentModel(softDoc.model, measurer);
		const firstPageBlock = at(at(at(laid.pages, 0).columns, 0).blocks, 0);
		expect(firstPageBlock.kind).toBe('paragraph');
		if (firstPageBlock.kind === 'paragraph') {
			expect(lineText(at(firstPageBlock.lines, 0).fragments)).toBe('cooperate');
		}
	});
});
