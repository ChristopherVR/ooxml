// U+FE0F (emoji presentation selector) takes no width and draws nothing (issue #439).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';
import { createFontFileMeasurer } from './font-measurer';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
// Glyph 1 (600 units) is the emoji, glyph 2 (700 units) maps U+FE0F as Segoe UI Emoji does.
const FONT = syntheticTrueType([
	[0x270d, 1],
	[0xfe0f, 2],
]);
const fonts = [{ family: 'Emoji Test', data: FONT }];

async function streamOf(text: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Emoji Test" w:hAnsi="Emoji Test"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Emoji Test" w:hAnsi="Emoji Test"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> A</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const measurer = createFontFileMeasurer(fonts);
	const { bytes } = await exportDocxPdf(model, { measurer, fonts });
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

/** The x of the text that follows the emoji (the glyph run after the first one). */
function nextTextX(stream: string): number {
	console.log(stream);
	const xs = [...stream.matchAll(/ ([0-9.]+) [0-9.]+ Tm/g)].map((m) => Number(m[1]));
	return xs[xs.length - 1]!;
}

describe('variation selector 16', () => {
	it('measures as its base character alone', () => {
		const measurer = createFontFileMeasurer(fonts);
		const font = { family: 'Emoji Test', sizePx: 100 };
		expect(measurer.widthOf('✍️', font)).toBe(measurer.widthOf('✍', font));
		expect(measurer.widthOf('✍', font)).toBeCloseTo(60, 5);
	});

	it('leaves no gap after the emoji in the exported text', async () => {
		const bug = await streamOf('✍️');
		const control = await streamOf('✍');
		expect(nextTextX(bug)).toBeCloseTo(nextTextX(control), 3);
	});
});
