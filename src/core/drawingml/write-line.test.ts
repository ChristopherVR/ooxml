import { describe, expect, it } from 'vitest';
import { NS, buildXml, parseXml } from '../xml/index';
import { parseDrawingLine } from './drawing-fill';
import type { DrawingFill } from './types';
import { drawingLineXml, patchDrawingLine } from './write-line';

const solid = (value: string): DrawingFill => ({
	kind: 'solid',
	color: { kind: 'srgb', value, transforms: [] },
});

describe('drawingLineXml', () => {
	it('writes width, cap, fill and dash in schema order and reads back', () => {
		const line = { widthEmu: 12700, cap: 'rnd', fill: solid('FF0000'), dash: 'dash' };
		const xml = drawingLineXml(line);
		expect(xml).toBe(
			`<a:ln xmlns:a="${NS.a}" w="12700" cap="rnd"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln>`,
		);
		expect(parseDrawingLine(parseXml(xml).documentElement)).toEqual(line);
	});
});

describe('patchDrawingLine', () => {
	it('keeps the join, compound type and arrow heads it does not model', () => {
		const doc = parseXml(
			`<a:ln xmlns:a="${NS.a}" w="9525" cmpd="sng"><a:noFill/><a:prstDash val="sysDot"/><a:round/><a:headEnd type="arrow"/></a:ln>`,
		);
		patchDrawingLine(doc.documentElement, { widthEmu: 19050, fill: solid('00FF00') });
		expect(buildXml(doc)).toBe(
			`<a:ln xmlns:a="${NS.a}" w="19050" cmpd="sng"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:round/><a:headEnd type="arrow"/></a:ln>`,
		);
	});

	it('leaves a fill the shared writer cannot express in place', () => {
		const doc = parseXml(`<a:ln xmlns:a="${NS.a}"><a:pattFill prst="pct5"/></a:ln>`);
		patchDrawingLine(doc.documentElement, {
			widthEmu: 100,
			fill: { kind: 'pattern', preset: 'pct5' },
		});
		expect(buildXml(doc)).toBe(`<a:ln xmlns:a="${NS.a}" w="100"><a:pattFill prst="pct5"/></a:ln>`);
	});
});
