// A colour glyph (COLR/CPAL) is drawn as its coloured layers, as Word does (issue #437).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { parseTrueType } from '../../text/truetype/parse';
import { colourLayers } from '../../text/truetype/colr';
import { colourTrueType } from '../../text/truetype/test-support/colour-font';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';
import { createFontFileMeasurer } from './font-measurer';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function streamOf(text: string, data: Uint8Array, color = '') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Emoji Test" w:hAnsi="Emoji Test"/>${color}<w:sz w:val="36"/></w:rPr><w:t>${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const fonts = [{ family: 'Emoji Test', data }];
	const { bytes } = await exportDocxPdf(model, { measurer: createFontFileMeasurer(fonts), fonts });
	const raw = Buffer.from(bytes).toString('latin1');
	const out: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			out.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return out.join('\n');
}

/** The page content: its colour operators and text operators in order. */
const operators = (stream: string) =>
	[
		...stream.matchAll(
			/^([\d. ]+) rg$|^BT .* ([\d.]+) ([\d.]+) Tm .*\[(<[^\]]*>)[^\]]*\] TJ ET$/gm,
		),
	].map((m) => (m[1] ? `rg ${m[1]}` : `Tm ${m[2]} ${m[4]}`));

describe('colour glyphs', () => {
	it('reads the layers of the COLR base glyph records with the first CPAL palette', () => {
		const layers = colourLayers(parseTrueType(colourTrueType())!)!;
		expect(layers(1)).toEqual([
			{ glyph: 2, color: [1, 0, 0] },
			{ glyph: 3, color: [0, 0, 1] },
			{ glyph: 2, color: undefined },
		]);
		expect(layers(2)).toBeUndefined();
		expect(colourLayers(parseTrueType(syntheticTrueType())!)).toBeUndefined();
	});

	it('draws a colour glyph as its layers in their palette colours, not as the base outline', async () => {
		const stream = await streamOf('☀', colourTrueType(), '<w:color w:val="00FF00"/>');
		const ops = operators(stream);
		// Layer 1 red, layer 2 blue, layer 3 in the text colour, then the text colour again.
		const colours = ops.filter((op) => op.startsWith('rg')).slice(-4);
		expect(colours).toEqual(['rg 1 0 0', 'rg 0 0 1', 'rg 0 1 0', 'rg 0 1 0']);
		const texts = ops.filter((op) => op.startsWith('Tm'));
		expect(texts).toHaveLength(3);
		// The layers sit at the glyph's origin and are distinct glyphs (layer 1 and 3 share glyph 2).
		expect(new Set(texts.map((text) => text.split(' ')[1])).size).toBe(1);
		expect(texts[0]!.split(' ')[2]).toBe(texts[2]!.split(' ')[2]);
		expect(texts[1]!.split(' ')[2]).not.toBe(texts[0]!.split(' ')[2]);
	});

	it('leaves text without colour glyphs as one text operator', async () => {
		const stream = await streamOf('B', colourTrueType());
		expect(operators(stream).filter((op) => op.startsWith('Tm'))).toHaveLength(1);
	});

	it('draws the glyph as an ordinary black outline in a font without colour tables', async () => {
		const plain = syntheticTrueType([[0x2600, 1]]);
		const stream = await streamOf('☀', plain);
		expect(operators(stream).filter((op) => op.startsWith('Tm'))).toHaveLength(1);
	});
});
