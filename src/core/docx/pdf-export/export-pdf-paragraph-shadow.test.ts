// Paragraph border shadows (`w:shadow`, issue 377) in the PDF: the right and bottom lines are
// drawn again just outside the box, and the box itself keeps the geometry it has without a shadow.
import JSZip from 'jszip';
import { inflateSync } from 'node:zlib';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const side = (name: string, space: number, shadow: boolean) =>
	`<w:${name} w:val="single" w:sz="12" w:space="${space}" w:color="auto"${shadow ? ' w:shadow="1"' : ''}/>`;
const box = (shadow: boolean) =>
	`<w:pBdr>${side('top', 1, shadow)}${side('left', 4, shadow)}${side('bottom', 1, shadow)}${side('right', 4, shadow)}</w:pBdr>`;

/** The filled rectangles of the page content: `[x, y, width, height]` in PDF points. */
async function rects(shadow: boolean): Promise<number[][]> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:pPr>${box(shadow)}<w:jc w:val="center"/></w:pPr><w:r><w:t>Bordered paragraph</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer() });
	const raw = Buffer.from(bytes).toString('latin1');
	const found: number[][] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (!/\/Filter \/FlateDecode/.test(match[1]!) || /\/Subtype \/Image|\/Length1/.test(match[1]!))
			continue;
		const stream = inflateSync(bytes.subarray(start, start + length)).toString('latin1');
		for (const rect of stream.matchAll(/^([\d.-]+) ([\d.-]+) ([\d.-]+) ([\d.-]+) re$/gm))
			found.push(rect.slice(1, 5).map(Number));
	}
	return found.filter(([, , width, height]) => width! < 790 && height! < 500);
}

it('draws the shadow of the right and bottom lines outside the box', async () => {
	const plain = await rects(false);
	const shadowed = await rects(true);
	const line = (1.5 * 96) / 72;
	expect(shadowed).toHaveLength(plain.length + 2);
	const lines = (all: number[][]) => ({
		right: all.filter(([, , width, height]) => width === line && height! > 3 * line),
		bottom: all.filter(([, , width, height]) => height === line && width! > 3 * line),
	});
	const before = lines(plain);
	const after = lines(shadowed);
	const boxRight = Math.max(...before.right.map(([x]) => x!));
	const boxBottom = Math.max(...before.bottom.map(([, y]) => y!));
	// The shadow strips start where the box's right and bottom lines end.
	expect(after.right.map(([x]) => x!)).toContain(boxRight + line);
	expect(after.bottom.map(([, y]) => y!).some((y) => y > boxBottom)).toBe(true);
	const shadowRight = after.right.find(([x]) => x! > boxRight)!;
	const shadowBottom = after.bottom.find(([, y]) => y! > boxBottom)!;
	expect(shadowRight[2]).toBeCloseTo(line, 5);
	expect(shadowBottom[0]! + shadowBottom[2]!).toBeCloseTo(boxRight + 2 * line, 5);
});
