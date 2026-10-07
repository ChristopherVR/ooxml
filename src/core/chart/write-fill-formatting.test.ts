import { expect, it } from 'vitest';
import { NS, buildXml, parseXml } from '../xml';
import { readChartFormatting } from './read-formatting';
import { writeChartFillFormatting } from './write-fill-formatting';
import { setDrawingFillXml } from '../diagram/write-fill';

it('patches fills without replacing outlines, effects, transforms or unknown extensions', () => {
	const doc = parseXml(
		`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}"><c:chart><c:plotArea><c:layout/><c:barChart/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"/><a:effectDag/><a:extLst><a:ext uri="keep"/></a:extLst></c:spPr></c:plotArea></c:chart><c:spPr><a:xfrm/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c:spPr><c:txPr/></c:chartSpace>`,
	);
	const formatting = readChartFormatting(doc.documentElement)!;
	expect(writeChartFillFormatting(doc.documentElement, formatting)).toBe(false);
	formatting.entries.plotArea!.fill = { kind: 'none' };
	formatting.entries.chartArea!.fill = {
		kind: 'solid',
		color: { kind: 'scheme', value: 'accent2', transforms: [{ name: 'alpha', value: '50000' }] },
	};
	expect(writeChartFillFormatting(doc.documentElement, formatting)).toBe(true);
	const xml = buildXml(doc);
	expect(xml).toContain('<a:noFill/><a:ln w="12700"/><a:effectDag/><a:extLst>');
	expect(xml).toContain('uri="keep"');
	expect(xml).toContain('<a:xfrm/><a:solidFill>');
	expect(readChartFormatting(doc.documentElement)?.entries.chartArea?.fill).toEqual(
		formatting.entries.chartArea!.fill,
	);
	expect(
		writeChartFillFormatting(doc.documentElement, readChartFormatting(doc.documentElement)),
	).toBe(false);
});

it('inserts chart-space shape properties before text properties and retains unsupported fills', () => {
	const doc = parseXml(
		`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}"><c:chart/><c:txPr/><c:extLst/></c:chartSpace>`,
	);
	expect(
		writeChartFillFormatting(doc.documentElement, {
			sourceXml: '',
			entries: { chartArea: { sourceXml: '', fill: { kind: 'none' } } },
		}),
	).toBe(true);
	expect(buildXml(doc)).toContain('<c:spPr><a:noFill/></c:spPr><c:txPr/>');
	const before = buildXml(doc);
	expect(
		writeChartFillFormatting(doc.documentElement, {
			sourceXml: '',
			entries: { chartArea: { sourceXml: '', fill: { kind: 'unsupported', element: 'pattFill' } } },
		}),
	).toBe(false);
	expect(buildXml(doc)).toBe(before);
});

it('inserts a fill before effect containers and rejects malformed fragments before mutation', () => {
	const doc = parseXml(
		`<a:spPr xmlns:a="${NS.a}"><a:xfrm/><a:prstGeom prst="rect"/><a:effectDag/><a:extLst/></a:spPr>`,
	);
	setDrawingFillXml(doc.documentElement, '<a:noFill/>');
	expect(buildXml(doc)).toContain('<a:prstGeom prst="rect"/><a:noFill/><a:effectDag/>');
	const before = buildXml(doc);
	expect(() => setDrawingFillXml(doc.documentElement, '<a:noFill/><a:solidFill/>')).toThrow(
		'Invalid DrawingML fill',
	);
	expect(buildXml(doc)).toBe(before);
});
