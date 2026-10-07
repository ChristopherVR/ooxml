import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { insertRectangle } from './viewer-draw-tool.js';
import { renderPage } from './render-svg.js';
import { exportPageSvg } from './export-svg.js';
import { cell, fixture } from '../../../core/visio/test-fixtures.js';
import evidence from '../../../core/visio/__fixtures__/page-scales-native.json';

it('inserts, exports, undoes and redoes a rectangle at the pointer location on a scaled page', async () => {
	const bytes = await fixture({
		pages: [
			{ id: '0', pageCells: cell('DrawingScale', 2) + cell('PageScale', 1), contents: '<Shapes/>' },
		],
	});
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (source, edits) => {
			const saved = await editVsdx(source, edits);
			return { ...saved, document: await parseVsdx(saved.bytes) };
		},
	);
	try {
		await controller.load(bytes);
		const page = controller.state.document!.pages[0]!;
		const id = await insertRectangle(controller, page, { x: 1, y: 2 }, { width: 1.5, height: 1 });
		const created = controller.state.document!.pages[0]!.shapes[0]!;
		expect(created.id).toBe(id);
		expect([created.width, created.height]).toEqual([1.5, 1]);
		expect(created.transform.slice(4)).toEqual([0.25, page.height - 2.5]);
		expect((await parseVsdx(controller.exportVsdx().bytes)).pages[0]!.shapes[0]!.width).toBe(1.5);
		await controller.undo();
		expect(controller.state.document!.pages[0]!.shapes).toHaveLength(0);
		await controller.redo();
		expect(controller.state.document!.pages[0]!.shapes[0]!.width).toBe(1.5);
	} finally {
		controller.destroy();
	}
});

const native = process.env.VISIO_NATIVE_PAGE_SCALES_DIR;
describe.skipIf(!native)('native scaled page SVG rendering', () => {
	it('matches page dimensions, line stems, rectangle bounds and physical text placement', async () => {
		const model = await parseVsdx(await readFile(resolve(native!, 'page-scales.vsdx')));
		for (const row of evidence.cases) {
			const page = model.pages.find((p) => p.id === row.pageId)!;
			const result = renderPage(model, page);
			expect(Number(result.svg.getAttribute('viewBox')!.split(' ')[2])).toBeCloseTo(
				row.nativePaperWidth,
				12,
			);
			const line = result.svg.querySelector(`[data-shape-id="${row.lineId}"] path`)!;
			const actualEnd = Number(line.getAttribute('d')!.split(' ').at(-2));
			const nativeEnd = Number(row.nativeLinePath.match(/L([0-9.]+)/)![1]) / 72;
			expect(actualEnd).toBeCloseTo(nativeEnd, 3);
			expect(Number(line.getAttribute('stroke-width'))).toBe(row.nativeLineStroke);
			const text = result.svg.querySelector(`[data-shape-id="${row.rectangleId}"] text`)!;
			expect(Number(text.getAttribute('font-size'))).toBe(row.fontSize);
			const run = text.querySelector('tspan')!;
			expect(Number(run.getAttribute('x'))).toBeCloseTo(row.nativeTextX, 12);
			result.dispose();
			const exported = exportPageSvg(model, model.pages.indexOf(page));
			expect(exported.svg).toContain(`width="${row.nativePaperWidth}in"`);
		}
	}, 30_000);
});
