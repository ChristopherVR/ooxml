import { describe, expect, it } from 'vitest';

import { readDiagramDrawingShapes, themeColorResolver } from './drawing-resolve';

const NS =
	'xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const theme = { scheme: (name: string) => ({ accent1: '#4472C4', lt1: '#FFFFFF' })[name] };

const drawing = (shapes: string) =>
	`<dsp:drawing ${NS}><dsp:spTree>${shapes}</dsp:spTree></dsp:drawing>`;

describe('readDiagramDrawingShapes', () => {
	it('resolves geometry, fill, line, text style and the text frame into pixels and hex', () => {
		const { shapes, issues } = readDiagramDrawingShapes(
			drawing(`<dsp:sp modelId="{A}"><dsp:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="95250" y="190500"/><a:ext cx="952500" cy="476250"/></a:xfrm>
			<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
			<a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill>
			<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></dsp:spPr>
			<dsp:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>
			<dsp:txBody><a:bodyPr lIns="9525" anchor="ctr"/><a:p><a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1800" b="1"><a:latin typeface="Calibri"/></a:rPr><a:t>Alpha</a:t></a:r></a:p><a:p><a:r><a:t>Beta</a:t></a:r></a:p></dsp:txBody>
			<dsp:txXfrm><a:off x="95250" y="190500"/><a:ext cx="952500" cy="476250"/></dsp:txXfrm></dsp:sp>`),
			{ resolveColor: themeColorResolver(theme) },
		);
		expect(issues).toEqual([]);
		expect(shapes).toEqual([
			{
				id: '{A}',
				shapeType: 'roundRect',
				x: 10,
				y: 20,
				width: 100,
				height: 50,
				rotation: 90,
				flipHorizontal: true,
				shapeAdjustments: { adj: 10000 },
				fillColor: '#4472C4',
				fillOpacity: 0.5,
				strokeColor: '#FF0000',
				strokeWidth: 2,
				text: 'Alpha\nBeta',
				fontSize: 24,
				fontColor: '#FFFFFF',
				fontFamily: 'Calibri',
				fontWeight: 700,
				lineHeightRatio: 0.9,
				textInsetLeft: 1,
				textVerticalAnchor: 'ctr',
				textFrameX: 10,
				textFrameY: 20,
				textFrameWidth: 100,
				textFrameHeight: 50,
			},
		]);
	});

	it('keeps zero-height lines and drops other frames thinner than half a pixel', () => {
		const shape = (id: string, prst: string, cy: number) =>
			`<dsp:sp modelId="${id}"><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="${cy}"/></a:xfrm><a:prstGeom prst="${prst}"/><a:noFill/></dsp:spPr></dsp:sp>`;
		const { shapes } = readDiagramDrawingShapes(
			drawing(shape('line', 'line', 0) + shape('hair', 'rect', 71) + shape('box', 'rect', 9525)),
			{ resolveColor: themeColorResolver(theme) },
		);
		expect(shapes.map((view) => [view.id, view.fillNone])).toEqual([
			['line', true],
			['box', true],
		]);
	});
});
