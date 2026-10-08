import { expect, it } from 'vitest';
import { NS, elements, first, parseXml } from '../xml/index';
import { parseDrawingFill } from './drawing-fill';
import { drawingFillXml } from './write-fill';

it('patches gradient stops while retaining source flags and extensions', () => {
	const fill = parseDrawingFill(
		parseXml(
			`<a:spPr xmlns:a="${NS.a}"><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="123456"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/><a:tileRect l="100"/><a:extLst/></a:gradFill></a:spPr>`,
		).documentElement,
	)!;
	if (fill.kind !== 'gradient') throw new Error('Expected gradient');
	expect(fill.scaled).toBe(false);
	fill.stops[0]!.color.value = 'ABCDEF';
	const root = parseXml(drawingFillXml(fill)!).documentElement;
	expect(root.getAttribute('rotWithShape')).toBe('1');
	expect(first(root, 'lin', NS.a)?.getAttribute('scaled')).toBe('0');
	fill.scaled = true;
	expect(
		first(parseXml(drawingFillXml(fill)!).documentElement, 'lin', NS.a)?.getAttribute('scaled'),
	).toBe('1');
	expect(first(root, 'tileRect', NS.a)?.getAttribute('l')).toBe('100');
	expect(first(root, 'extLst', NS.a)).toBeDefined();
	expect(root.getElementsByTagNameNS(NS.a, 'srgbClr')[0]?.getAttribute('val')).toBe('ABCDEF');
	delete fill.angle;
	fill.path = 'circle';
	fill.fillToRect = { l: 0.2, t: 0.3, r: 0.4, b: 0.5 };
	const radial = parseXml(drawingFillXml(fill)!).documentElement;
	expect(first(radial, 'lin', NS.a)).toBeUndefined();
	expect(first(first(radial, 'path', NS.a), 'fillToRect', NS.a)?.getAttribute('l')).toBe('20000');
	expect(elements(radial).map((node) => node.localName)).toEqual([
		'gsLst',
		'path',
		'tileRect',
		'extLst',
	]);
	delete fill.fillToRect;
	expect(drawingFillXml(fill)).not.toContain('fillToRect');
});

it('writes explicit no-fill and reports unsupported fill authoring', () => {
	expect(drawingFillXml({ kind: 'none' })).toBe('<a:noFill/>');
	expect(drawingFillXml({ kind: 'unsupported', element: 'grpFill' })).toBeUndefined();
});

it('retains unmodelled native gradient stops when editing recognized stops', () => {
	const fill = parseDrawingFill(
		parseXml(
			`<a:spPr xmlns:a="${NS.a}"><a:gradFill><a:gsLst><a:gs pos="0"><a:unknownClr val="opaque"/></a:gs><a:gs pos="100000"><a:srgbClr val="123456"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:spPr>`,
		).documentElement,
	)!;
	if (fill.kind !== 'gradient') throw new Error('Expected gradient');
	fill.stops[0]!.color.value = 'ABCDEF';
	const xml = drawingFillXml(fill)!;
	expect(xml).toContain('unknownClr val="opaque"');
	expect(parseXml(xml).getElementsByTagNameNS(NS.a, 'srgbClr')[0]?.getAttribute('val')).toBe(
		'ABCDEF',
	);
});
