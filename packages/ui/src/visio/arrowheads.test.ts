import { describe, expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';
import { demoDocument } from './demo-document.js';
import { renderPage } from './render-svg.js';
import { exportPageSvg } from './export-svg.js';

function scene(path = 'M 0 0 L 1 0') {
	const model = structuredClone(demoDocument);
	const shape = model.pages[0]!.shapes[1]!;
	shape.geometry = [{ path, fill: false, stroke: true }];
	shape.style.startArrow = shape.style.endArrow = 5;
	model.pages[0]!.shapes = [shape];
	return { model, shape };
}

// Independent analytic checks of the glyph's closed, inward-curved-base topology.
// These are SVG contract checks, not browser pixels or a Microsoft Visio oracle.
function checkConcaveMarker(marker: Element) {
	const glyph = marker.querySelector('path')!;
	expect(glyph.getAttribute('stroke-linejoin')).toBe('round');
	const tokens = glyph.getAttribute('d')!.split(/\s+/);
	expect(tokens.filter((v) => /^[A-Z]$/.test(v))).toEqual(['M', 'L', 'L', 'Q', 'Z']);
	const values = tokens.filter((v) => !/^[A-Z]$/.test(v)).map(Number);
	expect(values.every(Number.isFinite)).toBe(true);
	const [baseX, top, tipX, tipY, bottomX, bottom, controlX, controlY, endX, endY] = values as [
		number,
		number,
		number,
		number,
		number,
		number,
		number,
		number,
		number,
		number,
	];
	expect(bottomX).toBe(baseX);
	expect([endX, endY]).toEqual([baseX, top]);
	expect(tipX).toBeGreaterThan(controlX);
	expect(controlX).toBeGreaterThan(baseX);
	expect(top).toBeLessThan(tipY);
	expect(bottom).toBeGreaterThan(tipY);
	// Quadratic midpoint lies inside the straight base-to-tip triangle.
	const midpointX = (bottomX + 2 * controlX + endX) / 4;
	const midpointY = (bottom + 2 * controlY + endY) / 4;
	expect(midpointX).toBeGreaterThan(baseX);
	expect(midpointX).toBeLessThan(tipX);
	expect(midpointY).toBe(tipY);
	const [x, y, w, h] = marker.getAttribute('viewBox')!.split(' ').map(Number) as [
		number,
		number,
		number,
		number,
	];
	// The control hull and half-unit outline fit wholly within the viewport.
	for (let i = 0; i < values.length; i += 2) {
		expect(values[i]! - 0.5).toBeGreaterThanOrEqual(x);
		expect(values[i]! + 0.5).toBeLessThanOrEqual(x + w);
		expect(values[i + 1]! - 0.5).toBeGreaterThanOrEqual(y);
		expect(values[i + 1]! + 0.5).toBeLessThanOrEqual(y + h);
	}
	expect(Number(marker.getAttribute('refX'))).toBe(tipX);
	expect(Number(marker.getAttribute('refY'))).toBe(tipY);
	expect(marker.getAttribute('orient')).toBe('auto-start-reverse');
}

describe('code-5 concave arrows', () => {
	it.each(['M 0 0 L 1 0', 'M 0 0 L 0 1', 'M 0 0 L -1 1', 'M 0 0 L 1 -1', 'M 0 0 L 0.000001 0'])(
		'keeps both endpoint anchors and local orientation for %s',
		(path) => {
			const { model } = scene(path);
			const result = renderPage(model, model.pages[0]!);
			const line = result.svg.querySelector('[data-shape-id="c1"] path')!;
			expect(line.getAttribute('d')).toBe(path);
			const markers = [...result.svg.querySelectorAll('marker')];
			expect(markers).toHaveLength(2);
			expect(new Set(markers.map((m) => m.id)).size).toBe(2);
			markers.forEach(checkConcaveMarker);
			for (const [i, side] of ['start', 'end'].entries())
				expect(line.getAttribute(`marker-${side}`)).toBe(`url(#${markers[i]!.id})`);
			expect(result.warnings).not.toContain('Arrowhead style 5 is not rendered in this build.');
			expect(result.warnings.some((w) => w.includes('sizing is approximate'))).toBe(true);
			result.dispose();
		},
	);
	it.each([0, 0.35, 1])('uses safe paint and opacity %s once on the glyph', (opacity) => {
		const { model, shape } = scene();
		shape.style.lineColor = 'url(https://invalid.example/paint)';
		shape.style.lineOpacity = opacity;
		const result = renderPage(model, model.pages[0]!);
		for (const glyph of result.svg.querySelectorAll('marker path')) {
			expect(glyph.getAttribute('fill')).toBe('#273c46');
			expect(glyph.getAttribute('stroke')).toBe('#273c46');
			expect(glyph.getAttribute('opacity')).toBe(String(opacity));
			expect(glyph.parentElement!.hasAttribute('opacity')).toBe(false);
		}
		result.dispose();
	});
	it.each(['NoLine', 'transparent-pattern'])('suppresses code 5 for %s', (mode) => {
		const { model, shape } = scene();
		if (mode === 'NoLine') shape.geometry[0]!.stroke = false;
		else shape.style.linePattern = 0;
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelectorAll('marker,[marker-start],[marker-end]')).toHaveLength(0);
		result.dispose();
	});
	it.each([6, 17, 45, 254])('keeps unsupported code %s explicit', (code) => {
		const { model, shape } = scene();
		shape.style.startArrow = 0;
		shape.style.endArrow = code;
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelectorAll('marker')).toHaveLength(0);
		expect(result.warnings).toContain(`Arrowhead style ${code} is not rendered in this build.`);
		result.dispose();
	});
	it.each([0, 2, 6])('bounds size %s and exports inert local references', (size) => {
		const { model, shape } = scene();
		shape.style.startArrowSize = shape.style.endArrowSize = size;
		const result = exportPageSvg(model, 0);
		const xml = new DOMParser().parseFromString(result.svg, 'image/svg+xml');
		expect(xml.querySelector('parsererror,script,foreignObject')).toBeNull();
		for (const marker of xml.querySelectorAll('marker')) {
			checkConcaveMarker(marker);
			expect(Number(marker.getAttribute('markerWidth'))).toBeGreaterThan(0);
			expect(Number(marker.getAttribute('markerWidth'))).toBeLessThanOrEqual(0.37);
		}
		expect(xml.querySelectorAll('[marker-start],[marker-end]')).toHaveLength(1);
	});
});

const corpus = process.env.VISIO_ARROW_CORPUS_DIR;
describe.skipIf(!corpus)('hash-pinned real code-5 arrow corpus', () => {
	it('renders eight saved ends on 60973 page ID 0, without redistributing the source', async () => {
		const bytes = await readFile(resolve(corpus!, '60973.vsdx'));
		expect(createHash('sha256').update(bytes).digest('hex')).toBe(
			'c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03',
		);
		const model = await parseVsdx(bytes),
			page = model.pages.find((p) => p.id === '0')!;
		const result = renderPage(model, page);
		for (const id of ['3', '7', '8', '10']) {
			const shape = page.shapes.find((s) => s.id === id)!;
			expect([shape.style.startArrow, shape.style.endArrow]).toEqual([5, 5]);
			expect(shape.geometry[0]).toMatchObject({ fill: false, stroke: true });
			const group = result.svg.querySelector(`[data-shape-id="${id}"]`)!;
			expect(group.getAttribute('transform')).toBe(`matrix(${shape.transform.join(' ')})`);
			const path = group.querySelector('path')!;
			for (const side of ['start', 'end'])
				expect(path.getAttribute(`marker-${side}`)).toMatch(/^url\(#visio-arrow-\d+\)$/);
		}
		expect(result.svg.querySelectorAll('marker')).toHaveLength(8);
		result.svg.querySelectorAll('marker').forEach(checkConcaveMarker);
		result.dispose();
		const exported = exportPageSvg(model, model.pages.indexOf(page));
		const xml = new DOMParser().parseFromString(exported.svg, 'image/svg+xml');
		expect(xml.querySelectorAll('marker')).toHaveLength(8);
		expect(xml.querySelectorAll('[marker-start],[marker-end]')).toHaveLength(4);
		expect(xml.querySelector('parsererror,script,foreignObject')).toBeNull();
	}, 30_000);
});
