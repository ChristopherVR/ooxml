import { describe, expect, it } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { renderPage } from './render-svg.js';
import { exportPageSvg } from './export-svg.js';
import evidence from '../../../core/visio/__fixtures__/open-arrows-native.json';

describe('native open-arrow glyph coordinates', () => {
	it.each(evidence.cases)('matches code $code size $size weight $lineWidth', (row) => {
		const model = structuredClone(demoDocument);
		const page = model.pages[0]!;
		const shape = page.shapes[1]!;
		page.shapes = [shape];
		shape.style.startArrow = 0;
		shape.style.endArrow = row.code;
		shape.style.endArrowSize = row.size;
		shape.style.lineWidth = row.lineWidth;
		const result = renderPage(model, page);
		const glyph = result.svg.querySelector('marker path')!;
		const coordinates = (path: string) => path.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
		const native = coordinates(row.nativeGlyph);
		const expected = native.map((n, i) => n * row.extent * (i % 2 ? 1 : -1));
		const actual = coordinates(glyph.getAttribute('d')!);
		const pairs = (values: number[]) =>
			Array.from({ length: values.length / 2 }, (_, i) => [
				values[2 * i]!,
				values[2 * i + 1]!,
			]).sort((a, b) => a[0]! - b[0]! || a[1]! - b[1]!);
		const sorted = pairs(actual);
		pairs(expected).forEach((p, i) => {
			expect(sorted[i]![0]).toBeCloseTo(p[0]!, 12);
			expect(sorted[i]![1]).toBeCloseTo(p[1]!, 12);
		});
		expect(glyph.getAttribute('stroke-width')).toBe(String(row.lineWidth));
		expect(glyph.getAttribute('fill')).toBe('none');
		result.dispose();
	});
});

describe('code-9 native tick markers', () => {
	it.each(['M 0 0 L 1 0', 'M 0 0 L 0 1', 'M 0 0 L -1 -1'])(
		'anchors both ticks without shortening %s',
		(path) => {
			const model = structuredClone(demoDocument);
			const page = model.pages[0]!;
			const shape = page.shapes[1]!;
			page.shapes = [shape];
			shape.geometry = [{ path, fill: false, stroke: true }];
			shape.style.startArrow = shape.style.endArrow = 9;
			shape.style.startArrowSize = 0;
			shape.style.endArrowSize = 6;
			shape.style.lineWidth = 0.01;
			shape.style.lineOpacity = 0.35;
			const result = renderPage(model, page);
			const markers = [...result.svg.querySelectorAll('marker')];
			expect(markers).toHaveLength(2);
			expect(markers.map((m) => m.querySelector('path')!.getAttribute('d'))).toEqual([
				'M -0.03 -0.03 L 0.03 0.03',
				'M -0.26 -0.26 L 0.26 0.26',
			]);
			for (const marker of markers) {
				expect(marker.getAttribute('refX')).toBe('0');
				expect(marker.getAttribute('refY')).toBe('0');
				expect(marker.getAttribute('orient')).toBe('auto-start-reverse');
				expect(marker.querySelector('path')!.getAttribute('stroke-width')).toBe('0.01');
				expect(marker.querySelector('path')!.getAttribute('stroke-linecap')).toBe('round');
				expect(marker.querySelector('path')!.getAttribute('opacity')).toBe('0.35');
			}
			expect(result.svg.querySelector('[data-shape-id="c1"] path')!.getAttribute('d')).toBe(path);
			expect(result.warnings.some((w) => /Arrowhead|arrowhead/.test(w))).toBe(false);
			result.dispose();
			const exported = exportPageSvg(model, 0);
			const xml = new DOMParser().parseFromString(exported.svg, 'image/svg+xml');
			expect(xml.querySelectorAll('marker')).toHaveLength(2);
			expect(xml.querySelector('parsererror,script,foreignObject')).toBeNull();
		},
	);
});
