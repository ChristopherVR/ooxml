import { expect, it } from 'vitest';
import { NS, first, parseXml } from '../xml/index';
import { resolveDrawingShadow, svgDropShadowElement } from './drawing-shadow';

it('uses native defaults while resolving themed shadow alpha', () => {
	const root = parseXml(
		`<a:effectLst xmlns:a="${NS.a}"><a:outerShdw><a:schemeClr val="dk1"><a:alpha val="25000"/></a:schemeClr></a:outerShdw></a:effectLst>`,
	).documentElement;
	expect(resolveDrawingShadow(root, { scheme: () => '#000000' })).toEqual({
		color: '#000000',
		opacity: 0.25,
		blur: 0,
		dx: 0,
		dy: 0,
	});
	first(root, 'outerShdw', NS.a)!.setAttribute('sy', '23000');
	expect(resolveDrawingShadow(root, { scheme: () => '#000000' })).toBeUndefined();
});

it('rejects invalid geometry instead of emitting nonfinite SVG filters', () => {
	for (const value of ['NaN', '-1']) {
		const root = parseXml(
			`<a:effectLst xmlns:a="${NS.a}"><a:outerShdw blurRad="${value}"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>`,
		).documentElement;
		expect(resolveDrawingShadow(root, { scheme: () => undefined })).toBeUndefined();
	}
	expect(
		svgDropShadowElement(
			{ color: '#000000', opacity: 0.63, blur: 6, dx: 0, dy: 2 },
			{ number: String, color: (value) => value },
		),
	).toContain('stdDeviation="3"');
});
