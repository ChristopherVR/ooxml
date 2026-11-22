import { expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { renderSvgPreview } from './svg-preview';
import { renderWordArtSvg } from './word-art-svg';

it('keeps preview dimensions, local paints, filters and text', () => {
	const svg = parseXml(
		renderSvgPreview(
			32,
			20,
			'<linearGradient id="paint"><stop stop-color="#ABCDEF"/></linearGradient>',
			'<g style="filter:grayscale(100%) blur(1px)"><rect fill="url(#paint)"/><text>A &amp; B</text></g>',
			{ width: 64, height: 40 },
		),
	).documentElement;
	expect(svg.getAttribute('width')).toBe('32');
	expect(svg.getAttribute('height')).toBe('20');
	expect(svg.getAttribute('viewBox')).toBe('0 0 64 40');
	expect(svg.getElementsByTagName('linearGradient')[0]?.getAttribute('id')).toBe('paint');
	expect(svg.getElementsByTagName('rect')[0]?.getAttribute('fill')).toBe('url(#paint)');
	expect(svg.getElementsByTagName('g')[0]?.getAttribute('style')).toBe(
		'filter:grayscale(100%) blur(1px)',
	);
	expect(svg.getElementsByTagName('text')[0]?.textContent).toBe('A & B');
});

it('preserves WordArt text, gradients, outlines and local text paths', () => {
	const line = 'Sales & "Margin" <2026>';
	const svg = parseXml(
		renderWordArtSvg({
			id: 'Sales chart',
			style: {
				fill: '112233',
				gradient: { end: 'CCDDEE', angle: 0 },
				outline: '445566',
				shadow: '000000',
				warp: { preset: 'textArchUp', guides: [] },
			},
			lines: [line],
			width: 344,
			height: 80,
			fontPx: 32,
		}),
	).documentElement;
	expect(svg.getAttribute('data-word-art-transform')).toBe('textArchUp');
	expect(svg.getElementsByTagName('linearGradient')[0]?.getAttribute('id')).toBe(
		'Sales_chart-gradient',
	);
	const group = svg.getElementsByTagName('g')[0];
	expect(group?.getAttribute('fill')).toBe('url(#Sales_chart-gradient)');
	expect(group?.getAttribute('stroke')).toBe('#445566');
	expect(group?.getAttribute('style')).toContain('drop-shadow(');
	const path = svg.getElementsByTagName('path')[0];
	const textPath = svg.getElementsByTagName('textPath')[0];
	expect(path?.getAttribute('d')).toBeTruthy();
	expect(textPath?.getAttribute('href')).toBe(`#${path?.getAttribute('id')}`);
	expect(textPath?.textContent).toBe(line);
});
