// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer, layoutDocumentModel } from 'ooxml-core/docx/layout';
import { drawingsFixture } from '../../../core/docx/test-support/drawings-fixture';
import {
	BOXES,
	textBoxLayoutFixture,
} from '../../../core/docx/test-support/text-box-layout-fixture';
import { decoratePages } from './print-header-footer';
import { renderPrintLayout } from './print-layout';
import { must } from './test-support';

// Print Layout draws a text box's lines where the page engine laid them out (`TextBoxLayout`):
// absolutely placed lines in a layer clipped to the shape's text rectangle, turned for vertical
// text, in the body and in headers alike.

async function printed() {
	const { model } = await loadDocx(await textBoxLayoutFixture());
	const result = layoutDocumentModel(model, createFakeMeasurer());
	const handle = renderPrintLayout(result, undefined, { model });
	const page = must(handle.element.querySelector<HTMLElement>('.dve-print-page'), 'page');
	const floats = [...page.querySelectorAll<HTMLElement>('.dve-print-float')];
	const byName = new Map(BOXES.map((box, index) => [box.name, must(floats[index], box.name)]));
	return { result, page, byName };
}
const lineTexts = (element: HTMLElement) =>
	[...element.querySelectorAll('.dve-print-text-box .dve-print-line')].map((line) =>
		line.textContent?.trim(),
	);

describe('text boxes in Print Layout', () => {
	it('draw the engine’s lines in a layer clipped to the text rectangle', async () => {
		const { result, page, byName } = await printed();
		expect(page.querySelector('.dve-text-box-content')).toBeNull();
		const wrapped = must(byName.get('Wrapped'), 'Wrapped');
		const layer = must(wrapped.querySelector<HTMLElement>('.dve-print-text-box'), 'layer');
		const layout = must(result.pages[0]?.floats?.[0]?.textBoxLayout, 'layout');
		expect(layer.style.left).toBe(`${layout.frame.xPx}px`);
		expect(layer.style.width).toBe(`${layout.frame.widthPx}px`);
		expect(layer.style.height).toBe(`${layout.frame.heightPx}px`);
		const lines = lineTexts(wrapped);
		expect(lines[0]).toBe('Wrapped word1 word2');
		expect(lines.length).toBe(
			layout.blocks.reduce(
				(sum, block) => sum + (block.kind === 'paragraph' ? block.lines.length : 0),
				0,
			),
		);
		// The box's lines belong to no editable story.
		expect(wrapped.querySelector('.dve-print-text-box [data-block-id]')).toBeNull();
	});

	it('continue a linked box’s text in the next box', async () => {
		const { byName } = await printed();
		const first = lineTexts(must(byName.get('Linked first'), 'first'));
		const second = lineTexts(must(byName.get('Linked second'), 'second'));
		expect(first.join(' ')).not.toContain('word9');
		expect(second[0]).toMatch(/^word9\b/);
	});

	it('turn vert text clockwise and vert270 text anticlockwise', async () => {
		const { byName } = await printed();
		const turn = (name: string) =>
			must(byName.get(name)?.querySelector<HTMLElement>('.dve-print-text-box-text'), name).style
				.transform;
		expect(turn('Vertical')).toMatch(/^translate\([\d.]+px(, 0(px)?)?\) rotate\(90deg\)$/);
		expect(turn('Vertical 270')).toMatch(/^translate\(0(px)?, [\d.]+px\) rotate\(-90deg\)$/);
		expect(turn('Wrapped')).toBe('');
	});

	it('draw an autofit shape at the height its text takes', async () => {
		const { result, byName } = await printed();
		const fitted = must(result.pages[0]?.floats?.[4], 'Autofit float');
		expect(fitted.heightPx).toBeGreaterThan(BOXES[4]!.height * 96);
		expect(parseFloat(byName.get('Autofit')?.style.height ?? '')).toBeCloseTo(fitted.heightPx, 2);
	});

	it('draw a header text box with the engine too', async () => {
		const { model } = await loadDocx(await drawingsFixture());
		const result = layoutDocumentModel(model, createFakeMeasurer());
		const handle = renderPrintLayout(result, undefined, { model });
		const sheets = [...handle.element.querySelectorAll<HTMLElement>('.dve-print-page')];
		decoratePages(model, result.pages, sheets);
		const header = must(
			[...must(sheets[0], 'first page').querySelectorAll<HTMLElement>('.dve-print-text-box')].find(
				(layer) => layer.textContent?.includes('First Partner'),
			),
			'header text box',
		);
		expect(header.querySelector('.dve-print-line')?.textContent).toBe('First Partner');
		expect(sheets[0]?.querySelector('.dve-text-box-content')).toBeNull();
	});
});
