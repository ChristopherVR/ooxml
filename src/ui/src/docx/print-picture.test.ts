// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import { pictureElement, sanitizeChartSvg } from './print-picture';

const object = (svg: string) => ({
	partName: '',
	contentType: '',
	widthPx: 100,
	heightPx: 50,
	chart: { svg, label: 'Chart' },
});

describe('pictureElement chart SVG sanitising', () => {
	it('drops scripts, foreignObject, event handlers and unsafe links', () => {
		const svg =
			'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" onload="alert(1)">' +
			'<script>alert(2)</script>' +
			'<foreignObject><div>x</div></foreignObject>' +
			'<rect width="10" height="10" onclick="alert(3)"/>' +
			'<a href="javascript:alert(4)"><text>bad</text></a>' +
			'<a xlink:href="https://example.com/x"><text>remote</text></a>' +
			'<use href="#shape1"/>' +
			'<image xlink:href="data:image/png;base64,iVBORw0KGgo="/>' +
			'</svg>';

		const element = pictureElement(object(svg), undefined);
		const root = element.querySelector('svg');

		expect(root).not.toBeNull();
		expect(root!.hasAttribute('onload')).toBe(false);
		expect(element.querySelector('script')).toBeNull();
		expect(element.querySelector('foreignObject')).toBeNull();
		expect(element.querySelector('rect')!.hasAttribute('onclick')).toBe(false);
		expect(element.querySelectorAll('a')[0]!.hasAttribute('href')).toBe(false);
		expect(element.querySelectorAll('a')[1]!.getAttribute('xlink:href')).toBeNull();
		expect(element.querySelector('use')!.getAttribute('href')).toBe('#shape1');
		expect(element.querySelector('image')!.getAttribute('xlink:href')).toMatch(/^data:image\/png/u);
		expect(element.innerHTML).not.toContain('alert(');
		expect(element.innerHTML).not.toContain('example.com');
	});

	it('sanitises an SVG element in place', () => {
		const doc = new DOMParser().parseFromString(
			'<svg xmlns="http://www.w3.org/2000/svg"><script>x</script><g onmouseover="x"/></svg>',
			'image/svg+xml',
		);
		const svg = doc.documentElement;

		sanitizeChartSvg(svg);

		expect(svg.getElementsByTagName('script')).toHaveLength(0);
		expect(svg.getElementsByTagName('g')[0]!.hasAttribute('onmouseover')).toBe(false);
	});
});
