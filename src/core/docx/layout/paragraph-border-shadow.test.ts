// `w:shadow` on paragraph borders (issue 377): Word moves the box down by the line width, draws the
// right and bottom lines again just outside it, and keeps that room before the next paragraph.
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx, type Paragraph } from '../index';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import { paragraphFrameBounds, type LayoutParagraphBox } from './result';
import { paragraphShadowRects } from './paragraph-shadow';
import { at } from './test-helpers';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const side = (name: string, space: number, shadow: boolean) =>
	`<w:${name} w:val="single" w:sz="12" w:space="${space}" w:color="auto"${shadow ? ' w:shadow="1"' : ''}/>`;
const box = (shadow: boolean) =>
	`<w:pBdr>${side('top', 1, shadow)}${side('left', 4, shadow)}${side('bottom', 1, shadow)}${side('right', 4, shadow)}</w:pBdr>`;

async function boxes(shadow: boolean) {
	const zip = new JSZip();
	const paragraph = (text: string, props = '') =>
		`<w:p><w:pPr>${props}<w:spacing w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${paragraph('Before.')}${paragraph('Bordered paragraph', box(shadow))}${paragraph('Next paragraph.')}<w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	return {
		model,
		blocks: at(at(pages, 0).columns, 0).blocks as LayoutParagraphBox[],
	};
}

it('reads the shadow flag of each paragraph border side', async () => {
	const { model } = await boxes(true);
	const borders = (model.blocks[1] as Paragraph).borders;
	expect(borders?.right?.shadow).toBe(true);
	expect(borders?.top?.shadow).toBe(true);
	expect((await boxes(false)).model.blocks[1]).not.toHaveProperty('borders.right.shadow');
});

it('adds the shadow width above and below the box and pushes the next paragraph down', async () => {
	const plain = (await boxes(false)).blocks;
	const shadowed = (await boxes(true)).blocks;
	const line = (1.5 * 96) / 72;
	expect(at(shadowed, 1).heightPx - at(plain, 1).heightPx).toBeCloseTo(2 * line, 5);
	expect(at(shadowed, 2).yPx - at(plain, 2).yPx).toBeCloseTo(2 * line, 5);
	const frame = at(shadowed, 1).frame!;
	// The box keeps its size; it starts a line width lower, with the shadow below it.
	const bounds = paragraphFrameBounds(frame, at(shadowed, 1).heightPx);
	const unshadowed = paragraphFrameBounds(at(plain, 1).frame!, at(plain, 1).heightPx);
	expect(bounds.topPx).toBeCloseTo(line, 5);
	expect(bounds.heightPx).toBeCloseTo(unshadowed.heightPx, 5);
	expect(at(shadowed, 1).lines[0]!.yPx - at(plain, 1).lines[0]!.yPx).toBeCloseTo(line, 5);
});

it('draws the right and bottom shadow outside the box, one line wide', async () => {
	const frame = at((await boxes(true)).blocks, 1).frame!;
	const { heightPx } = paragraphFrameBounds(frame, at((await boxes(true)).blocks, 1).heightPx);
	const line = (1.5 * 96) / 72;
	const [right, bottom] = paragraphShadowRects(frame, heightPx);
	expect(right).toMatchObject({ xPx: frame.widthPx, yPx: line, widthPx: line });
	expect(right!.yPx + right!.heightPx).toBeCloseTo(heightPx, 5);
	expect(bottom).toMatchObject({ xPx: line, yPx: heightPx, heightPx: line });
	expect(bottom!.xPx + bottom!.widthPx).toBeCloseTo(frame.widthPx + line, 5);
	expect(paragraphShadowRects(at((await boxes(false)).blocks, 1).frame!, heightPx)).toEqual([]);
});

it('starts the right shadow at the top only for the first paragraph of a border group', async () => {
	const zip = new JSZip();
	const grouped = (text: string) =>
		`<w:p><w:pPr>${box(true)}<w:spacing w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${grouped('One')}${grouped('Two')}${grouped('Three')}<w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	const blocks = at(at(pages, 0).columns, 0).blocks as LayoutParagraphBox[];
	const line = (1.5 * 96) / 72;
	const strips = blocks.map((block) => {
		const { heightPx } = paragraphFrameBounds(block.frame!, block.heightPx);
		return paragraphShadowRects(block.frame!, heightPx);
	});
	// The right strip: below the top line of the first paragraph, flush with the top of the rest.
	expect(strips.map((rects) => rects[0]!.yPx)).toEqual([line, 0, 0]);
	// Only the last paragraph has a bottom line, so only it has a bottom strip.
	expect(strips.map((rects) => rects.length)).toEqual([1, 1, 2]);
});
