import { expect, it } from 'vitest';
import { resolveDiagramDrawingStyles } from './drawing-theme';
import type { DiagramDrawing } from './types';
import type { ThemeFormatScheme } from '../drawingml/format-scheme';
import { parseTheme as parseWordTheme } from '../docx/theme';
import { parseTheme as parseExcelTheme } from '../xlsx/read/theme';
import { themeDrawing } from '../docx/ui/smartart-theme';

const color = {
	kind: 'scheme' as const,
	value: 'accent2',
	transforms: [{ name: 'alpha', value: '50000' }],
};
const scheme: ThemeFormatScheme = {
	fillStyles: [
		{ fill: { kind: 'solid', color: { kind: 'scheme', value: 'phClr', transforms: [] } } },
	],
	lineStyles: [{ widthEmu: 19050, fill: { kind: 'solid', color } }],
	effectStyles: [{ effects: { effects: [{ kind: 'outerShadow', color, distanceEmu: 9525 }] } }],
	backgroundFillStyles: [],
};

it('resolves inherited styles on a render copy without freezing authored colors', () => {
	const drawing: DiagramDrawing = {
		issues: [],
		shapes: [
			{
				modelId: 'a',
				geometry: 'rect',
				frame: { x: 0, y: 0, width: 952500, height: 952500 },
				has3d: false,
				style: { fill: { index: 1, color }, line: { index: 1 }, effect: { index: 1 } },
				line: { widthEmu: 38100 },
			},
		],
	};
	const original = JSON.stringify(drawing);
	const resolved = resolveDiagramDrawingStyles(drawing, scheme);
	expect(resolved.shapes[0]?.fill).toEqual({ kind: 'solid', color });
	expect(resolved.shapes[0]?.line).toMatchObject({
		widthEmu: 38100,
		fill: { kind: 'solid', color },
	});
	expect(resolved.shapes[0]?.effects?.effects[0]?.kind).toBe('outerShadow');
	expect(JSON.stringify(drawing)).toBe(original);
	expect(JSON.stringify(scheme)).toContain('phClr');
});

it('keeps explicit no-fill and an empty direct effect list over inherited styles', () => {
	const drawing: DiagramDrawing = {
		issues: [],
		shapes: [
			{
				modelId: 'a',
				geometry: 'rect',
				frame: { x: 0, y: 0, width: 1, height: 1 },
				has3d: false,
				style: { fill: { index: 1 }, effect: { index: 1 } },
				fill: { kind: 'none' },
				effects: { effects: [] },
			},
		],
	};
	expect(resolveDiagramDrawingStyles(drawing, scheme).shapes[0]).toMatchObject({
		fill: { kind: 'none' },
		effects: { effects: [] },
	});
});

it('retains shared style matrices in both product theme adapters', () => {
	const xml =
		'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme/><a:fontScheme/><a:fmtScheme name="Authored"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>';
	expect(parseWordTheme(xml).formatScheme?.fillStyles).toHaveLength(1);
	expect(parseExcelTheme(xml).formatScheme?.fillStyles).toHaveLength(1);
});

it('retains Word scheme-color alpha on the render copy', () => {
	const drawing: DiagramDrawing = {
		issues: [],
		shapes: [
			{
				modelId: 'a',
				geometry: 'rect',
				frame: { x: 0, y: 0, width: 1, height: 1 },
				has3d: false,
				fill: { kind: 'solid', color },
			},
		],
	};
	const result = themeDrawing(drawing, {
		colors: { accent2: 'FF0000' },
		colorMapping: {},
		fonts: { major: {}, minor: {} },
	});
	expect(result.drawing.shapes[0]?.fill).toEqual({
		kind: 'solid',
		color: { kind: 'srgb', value: 'FF0000', transforms: [{ name: 'alpha', value: '50000' }] },
	});
	expect(result.alphaDropped).toBe(false);
	expect(drawing.shapes[0]?.fill).toEqual({ kind: 'solid', color });
});
