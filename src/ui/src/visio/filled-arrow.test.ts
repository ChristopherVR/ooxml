import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import { renderPage } from './render-svg.js';
import { exportPageSvg } from './export-svg.js';

describe('filled native straight-line markers', () => {
	it.each([2, 4, 5, 6])('uses native asymmetric begin/end setback for code %s', (code) => {
		const model = structuredClone(demoDocument),
			page = model.pages[0]!,
			shape = page.shapes[1]!;
		page.shapes = [shape];
		shape.geometry = [{ path: 'M 0 0 L 2 0', fill: false, stroke: true }];
		shape.style.startArrow = shape.style.endArrow = code;
		shape.style.startArrowSize = shape.style.endArrowSize = 2;
		shape.style.lineWidth = 0.01;
		const result = renderPage(model, page);
		const markers = [...result.svg.querySelectorAll('marker')];
		expect(markers).toHaveLength(2);
		const setback = [0, 0, 1, 0, 2, 1.75, 2.25][code]! * 0.045;
		expect(Number(markers[0]!.getAttribute('refX'))).toBeCloseTo(-(setback - 0.005), 12);
		expect(Number(markers[1]!.getAttribute('refX'))).toBeCloseTo(-setback, 12);
		const values = result.svg
			.querySelector('[data-shape-id="c1"] path')!
			.getAttribute('d')!
			.match(/-?\d+(?:\.\d+)?/g)!
			.map(Number);
		expect(values[0]).toBeCloseTo(setback - 0.005, 12);
		expect(values[2]).toBeCloseTo(setback, 12);
		expect(values[4]).toBeCloseTo(2 - setback, 12);
		result.dispose();
	});
	it.each([2, 4, 5, 6])('renders code %s at the original end tip with no outline', (code) => {
		const model = structuredClone(demoDocument),
			page = model.pages[0]!,
			shape = page.shapes[1]!;
		page.shapes = [shape];
		shape.geometry = [{ path: 'M 0 0 L 2 0', fill: false, stroke: true }];
		shape.style.startArrow = 0;
		shape.style.endArrow = code;
		shape.style.startArrowSize = shape.style.endArrowSize = 2;
		shape.style.lineWidth = 0.01;
		shape.style.lineOpacity = 0.35;
		const result = renderPage(model, page);
		const markers = [...result.svg.querySelectorAll('marker')];
		expect(markers).toHaveLength(1);
		const setback = [0, 0, 1, 0, 2, 1.75, 2.25][code]! * 0.045;
		for (const marker of markers) {
			expect(Number(marker.getAttribute('refX'))).toBeCloseTo(-setback, 12);
			expect(marker.getAttribute('orient')).toBe('auto-start-reverse');
			expect(marker.querySelector('path')!.getAttribute('stroke')).toBe('none');
			expect(marker.querySelector('path')!.getAttribute('opacity')).toBe('0.35');
		}
		const stroke = result.svg.querySelector('[data-shape-id="c1"] path')!;
		const values = stroke
			.getAttribute('d')!
			.match(/-?\d+(?:\.\d+)?/g)!
			.map(Number);
		expect(values[0]).toBe(0);
		expect(values[2]).toBeCloseTo(2 - setback, 12);
		expect(shape.geometry[0]!.path).toBe('M 0 0 L 2 0');
		expect(result.warnings.some((w) => /arrowhead/i.test(w))).toBe(false);
		result.dispose();
		const xml = new DOMParser().parseFromString(exportPageSvg(model, 0).svg, 'image/svg+xml');
		expect(xml.querySelectorAll('marker')).toHaveLength(1);
		expect(xml.querySelector('parsererror,script,foreignObject')).toBeNull();
	});
	it.each(['M 0 0 L 0 0', 'M 0 0 Q 1 1 2 0'])(
		'reports unsupported code-6 setbacks for %s',
		(path) => {
			const model = structuredClone(demoDocument),
				page = model.pages[0]!,
				shape = page.shapes[1]!;
			page.shapes = [shape];
			shape.geometry = [{ path, fill: false, stroke: true }];
			shape.style.startArrow = 0;
			shape.style.endArrow = 6;
			const result = renderPage(model, page);
			expect(result.svg.querySelector('marker')).toBeNull();
			expect(result.warnings).toContain(
				'Arrowhead style 6 requires a straight round-capped line with non-overlapping filled setbacks.',
			);
			result.dispose();
		},
	);
});

const native = process.env.VISIO_NATIVE_FILLED_ARROWS_DIR;
describe.skipIf(!native)('native-authored filled-arrow package', () => {
	it('parses and renders all 84 reference cases', async () => {
		const model = await parseVsdx(await readFile(resolve(native!, 'filled-arrows.vsdx')));
		const page = model.pages[0]!;
		expect(page.shapes).toHaveLength(84);
		const result = renderPage(model, page);
		expect(result.svg.querySelectorAll('marker')).toHaveLength(84);
		expect(result.warnings.some((w) => /arrowhead/i.test(w))).toBe(false);
		result.dispose();
	}, 30_000);
});

const bothNative = process.env.VISIO_NATIVE_FILLED_BOTH_DIR;
describe.skipIf(!bothNative)('native-authored two-ended filled-arrow package', () => {
	it('renders all 168 ends without fallback warnings', async () => {
		const model = await parseVsdx(await readFile(resolve(bothNative!, 'filled-arrows.vsdx')));
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelectorAll('marker')).toHaveLength(168);
		expect(result.warnings.some((w) => /arrowhead/i.test(w))).toBe(false);
		result.dispose();
	}, 30_000);
});
