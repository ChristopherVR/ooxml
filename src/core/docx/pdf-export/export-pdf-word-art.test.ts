import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx, type InlineImage, type Paragraph } from '../index';
import { createFakeMeasurer } from '../layout/index';
import type { WordArtStyle } from '../../drawingml/index';
import { describe, expect, it } from 'vitest';
import { outlineTrueType } from '../../text/truetype/test-support/outline-font';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFakeMeasurer();
const ARIAL = [
	{ family: 'Arial', bold: true, data: outlineTrueType() },
	{ family: 'Arial', data: outlineTrueType() },
];

/** A document holding one WordArt text box, saved and read back as Word would open it. */
async function wordArtDocument(text: string, wordArt: WordArtStyle) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	const blank = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const image: InlineImage = {
		relId: '',
		partName: '',
		contentType: '',
		widthPx: 288,
		heightPx: 144,
		unsupported: 'Text box',
		textBoxText: [text],
		textBoxEditable: true,
		textBoxBorder: false,
		textBoxContent: { blocks: [], wordArt },
	};
	(blank.model.blocks[0] as Paragraph).runs.push({ text: '', image });
	return (await loadDocx(await blank.save())).model;
}

/** Page content streams, image dictionaries and the PDF source. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	const images: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Subtype \/Image/.test(match[1]!)) images.push(match[1]!);
		else if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return { raw, content: streams.join('\n'), images };
}

describe('exportDocxPdf WordArt', () => {
	it('bends the glyph outlines along the warp, with invisible text over them', async () => {
		const model = await wordArtDocument('ABBA', {
			fill: '1769A5',
			outline: '103D63',
			shadow: '45460D',
			warp: { preset: 'textArchUp', guides: [] },
		});
		const { bytes, report } = await exportDocxPdf(model, { measurer, fonts: ARIAL });
		const pdf = inspect(bytes);
		// The glyph edges are cut into short pieces and bent: dozens of points and curves, stroked
		// (1 pt outline, 4/3 px) and then filled.
		expect(pdf.content).toMatch(
			/0\.062745 0\.239216 0\.388235 RG\n1\.333333 w 0 J 0 j \[\] 0 d\n(?:[\d. -]+ [mlc]\n|h\n){60,}S\n/,
		);
		expect(pdf.content).toMatch(/0\.090196 0\.411765 0\.647059 rg\n[^]*?\nf\n/);
		// Each letter is written invisibly (render mode 3), turned along the arch.
		const invisible = [
			...pdf.content.matchAll(/^(\S+) (\S+) (\S+) (\S+) \S+ \S+ cm\nBT [^\n]* 3 Tr /gm),
		];
		expect(invisible).toHaveLength(4);
		const angles = invisible.map((m) => Math.atan2(Number(m[2]), Number(m[1])));
		expect(angles[0]!).toBeLessThan(-0.05);
		expect(angles[3]!).toBeGreaterThan(0.05);
		// The soft shadow is one picture whose soft mask is the blurred silhouette.
		expect(pdf.images.some((image) => /\/SMask \d+ 0 R/.test(image))).toBe(true);
		expect(pdf.content).toMatch(/\/GS\d+ gs\n[^\n]* cm\n\/Im\d+ Do/);
		expect(report.embeddedFonts).toEqual(['Arial']);
		expect(report.notes.join(' ')).not.toMatch(/WordArt/);
	});

	it('fills gradient WordArt through the glyph clip with a shading', async () => {
		const model = await wordArtDocument('ABC', {
			fill: '1769A5',
			gradient: { end: 'B7DDF5', angle: 5400000 },
		});
		const pdf = inspect((await exportDocxPdf(model, { measurer, fonts: ARIAL })).bytes);
		expect(pdf.content).toMatch(/\nW n\n[\d.-]+ 0 0 [\d.-]+ [\d.-]+ [\d.-]+ cm\n\/Sh\d+ sh/);
		expect(pdf.raw).toMatch(/\/ShadingType 2 [^]*?\/Coords \[0\.5 0 0\.5 1\]/);
	});

	it('places real text letter by letter along the warp without a font file', async () => {
		const model = await wordArtDocument('Wave', {
			fill: '1769A5',
			gradient: { end: 'B7DDF5', angle: 0 },
			outline: '103D63',
			shadow: '45460D',
			warp: { preset: 'textWave1', guides: [] },
		});
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		// Fill, stroke and clip (mode 6), each letter with its own turned matrix, then the gradient.
		const main = /BT \/F\d+ [\d.]+ Tf 6 Tr((?: [^\n]*? Tm \(.\) Tj)+) ET\n[^]*?\/Sh\d+ sh/.exec(
			pdf.content,
		);
		expect(main).not.toBeNull();
		expect([...main![1]!.matchAll(/\((.)\) Tj/g)].map((m) => m[1]).join('')).toBe('Wave');
		const turns = [...main![1]!.matchAll(/ (\S+) (\S+) \S+ \S+ \S+ \S+ Tm/g)].map((m) =>
			Math.atan2(Number(m[2]), Number(m[1])),
		);
		expect(new Set(turns.map((t) => t.toFixed(2))).size).toBeGreaterThan(1);
		// The shadow copies are marked so they never extract.
		expect(pdf.content).toContain('/Span <</ActualText ()>> BDC');
		expect(report.substitutedFonts).toContainEqual({ family: 'Arial', standardFont: 'Helvetica' });
		expect(report.notes.join(' ')).toMatch(/placed and turned along the warp/);
	});

	it('outlines and shadows the laid-out lines of unwarped WordArt', async () => {
		const model = await wordArtDocument('ABBA', {
			fill: '1769A5',
			outline: '103D63',
			shadow: '45460D',
		});
		const { bytes } = await exportDocxPdf(model, { measurer, fonts: ARIAL });
		const pdf = inspect(bytes);
		// Fill then stroke in one pass (render mode 2), 1 px wide in the outline colour.
		expect(pdf.content).toMatch(
			/0\.062745 0\.239216 0\.388235 RG\n1 w 0 J 0 j \[\] 0 d\n2 Tr\n[^]*?BT \/E\d+ [^\n]* TJ ET/,
		);
		// The shadow, blurred from the glyph outlines, is drawn first.
		const shadow = pdf.content.indexOf('/Im0 Do');
		expect(shadow).toBeGreaterThan(-1);
		expect(shadow).toBeLessThan(pdf.content.indexOf('2 Tr'));
		expect(pdf.images.some((image) => /\/SMask/.test(image))).toBe(true);
	});
});
