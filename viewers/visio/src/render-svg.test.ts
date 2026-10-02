import { describe, it, expect } from 'vitest';
import { renderPage, safeColor } from './render-svg.js';
import { demoDocument } from './demo-document.js';

describe('SVG renderer', () => {
	it('renders groups, paths and text without injecting document markup', () => {
		const model = structuredClone(demoDocument);
		const page = model.pages[0]!;
		page.shapes[0]!.text.plainText = '<script>alert(1)</script>';
		const { svg, warnings } = renderPage(model, page);
		expect(svg.querySelectorAll('path').length).toBeGreaterThan(4);
		expect(svg.querySelector('script')).toBeNull();
		expect(svg.textContent).toContain('<script>');
		expect(svg.querySelector('g')?.getAttribute('transform')).toBe('translate(0 7) scale(1 -1)');
		expect(warnings.some((warning) => warning.includes('Text metrics'))).toBe(true);
	});
	it('blocks paint-server URLs and unsafe CSS values', () => {
		expect(safeColor('url(https://example.com/pixel.svg#paint)')).toBe('#273c46');
		expect(safeColor('#aabbcc')).toBe('#aabbcc');
	});
	it('handles cyclic backgrounds without recursion and paints backgrounds first', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.backgroundPageId = '2';
		model.pages[1]!.backgroundPageId = '1';
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelectorAll('[data-shape-id="a1"]')).toHaveLength(1);
		expect(result.svg.querySelector('[data-shape-id]')?.getAttribute('data-shape-id')).toBe('a1');
	});
	it('skips hidden shapes and applies explicit line/fill settings', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.hidden = true;
		const { svg } = renderPage(model, model.pages[0]!);
		expect(svg.querySelector('[data-shape-id="s1"]')).toBeNull();
		expect(svg.querySelector('[data-shape-id="c1"] path')?.getAttribute('fill')).toBe('none');
	});
});

describe('line ends', () => {
	it.each(['round', 'butt', 'square'] as const)('uses the normalized %s cap', (lineCap) => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[1]!.style.lineCap = lineCap;
		const result = renderPage(model, model.pages[0]!);
		expect(
			result.svg.querySelector('[data-shape-id="c1"] path')?.getAttribute('stroke-linecap'),
		).toBe(lineCap);
		result.dispose();
	});
	it('renders common arrowheads with unique local IDs and warns on unsupported codes', () => {
		const model = structuredClone(demoDocument);
		const line = model.pages[0]!.shapes[1]!;
		line.style.startArrow = 1;
		line.style.endArrow = 4;
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelectorAll('marker')).toHaveLength(2);
		expect(
			result.svg.querySelector('[data-shape-id="c1"] path')?.getAttribute('marker-end'),
		).toMatch(/^url\(#visio-arrow-/);
		line.style.endArrow = 45;
		expect(renderPage(model, model.pages[0]!).warnings).toContain(
			'Arrowhead style 45 is not rendered in this build.',
		);
	});
});

describe('group paint ordering', () => {
	it.each([0, 1, 2] as const)('honors DisplayMode %s for the group shape and text', (mode) => {
		const model = structuredClone(demoDocument);
		const group = model.pages[0]!.shapes[0]!;
		group.kind = 'group';
		group.groupDisplayMode = mode;
		group.children = [{ ...model.pages[0]!.shapes[2]!, id: 'child' }];
		model.pages[0]!.shapes = [group];
		const result = renderPage(model, model.pages[0]!);
		const element = result.svg.querySelector('[data-shape-id="s1"]')!;
		const own = element.querySelector(':scope > [data-geometry]');
		const child = element.querySelector(':scope > [data-shape-id="child"]');
		expect(child).not.toBeNull();
		if (mode === 0) expect(own).toBeNull();
		else if (mode === 1)
			expect(Array.from(element.children).indexOf(own!)).toBeLessThan(
				Array.from(element.children).indexOf(child!),
			);
		else
			expect(Array.from(element.children).indexOf(own!)).toBeGreaterThan(
				Array.from(element.children).indexOf(child!),
			);
	});
});

describe('portable SVG serialization', () => {
	it('emits one namespace declaration and reparses as valid XML', () => {
		const result = renderPage(demoDocument, demoDocument.pages[0]!);
		const xml = new XMLSerializer().serializeToString(result.svg);
		expect(xml.match(/xmlns="http:\/\/www.w3.org\/2000\/svg"/g)).toHaveLength(1);
		expect(
			new DOMParser().parseFromString(xml, 'image/svg+xml').querySelector('parsererror'),
		).toBeNull();
		result.dispose();
	});
});

describe('text backdrops', () => {
	it('draws a bounded per-line background behind text runs', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.text.backgroundColor = '#fefefe';
		model.pages[0]!.shapes[0]!.text.backgroundOpacity = 0.4;
		const result = renderPage(model, model.pages[0]!);
		const backdrop = result.svg.querySelector('[data-shape-id="s1"] rect')!;
		expect(backdrop.getAttribute('fill')).toBe('#fefefe');
		expect(backdrop.getAttribute('fill-opacity')).toBe('0.4');
		expect(Number(backdrop.getAttribute('width'))).toBeGreaterThan(0);
	});
});

describe('non-displayed master alternatives', () => {
	it('does not create empty interactive shapes when all own geometry is hidden', () => {
		const model = structuredClone(demoDocument);
		const hidden = model.pages[0]!.shapes[0]!;
		hidden.geometry = [];
		hidden.text.plainText = '';
		hidden.text.runs = [];
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelector('[data-shape-id="s1"]')).toBeNull();
	});
	it('preserves a non-painting group transform around visible children', () => {
		const model = structuredClone(demoDocument);
		const group = model.pages[0]!.shapes[0]!;
		group.kind = 'group';
		group.groupDisplayMode = 0;
		group.children = [{ ...model.pages[0]!.shapes[2]!, id: 'visible-child' }];
		model.pages[0]!.shapes = [group];
		const result = renderPage(model, model.pages[0]!);
		expect(
			result.svg.querySelector('[data-shape-id="s1"] [data-shape-id="visible-child"]'),
		).not.toBeNull();
	});
});

describe('page identity validation', () => {
	it('rejects an outside page rather than bypassing validated frame dimensions', () => {
		expect(() => renderPage(demoDocument, { ...demoDocument.pages[0]!, width: Infinity })).toThrow(
			'does not belong',
		);
	});
});
