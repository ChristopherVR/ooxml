import { describe, expect, it } from 'vitest';
import { NS, buildXml, parseXml } from '../xml';
import { readChartFormatting } from './read-formatting';
import { writeChartAxisFormatting } from './write-axis-formatting';

const document = (axes: string) =>
	parseXml(
		`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}"><c:chart><c:plotArea>${axes}</c:plotArea></c:chart></c:chartSpace>`,
	);
const axis = (kind: string, id: number, flags = '') =>
	`<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>${flags}<c:axPos val="b"/><c:numFmt formatCode="0.00" sourceLinked="0"/><c:crossAx val="${3 - id}"/><c:extLst><c:ext uri="kept"/></c:extLst></c:${kind}>`;

describe('shared axis formatting writer', () => {
	it('preserves hidden axes and their original label position byte for byte', () => {
		const doc = document(
			axis('catAx', 1, '<c:delete val="true"/><c:tickLblPos val="high"/>') +
				axis('valAx', 2, '<c:delete val="false"/><c:tickLblPos val="low"/>'),
		);
		const before = buildXml(doc);
		expect(
			writeChartAxisFormatting(doc.documentElement, readChartFormatting(doc.documentElement)),
		).toBe(false);
		expect(buildXml(doc)).toBe(before);
	});

	it.each(['catAx', 'dateAx', 'valAx'])(
		'maps %s and inserts flags in schema order without replacing styling',
		(kind) => {
			const doc = document(axis(kind, 1) + axis('valAx', 2));
			expect(
				writeChartAxisFormatting(doc.documentElement, {
					sourceXml: '',
					entries: {
						categoryAxis: { sourceXml: '', axisVisible: false, labelsVisible: false },
						valueAxis: { sourceXml: '', axisVisible: true, labelsVisible: false },
					},
				}),
			).toBe(true);
			const xml = buildXml(doc);
			expect(xml).toContain('</c:scaling><c:delete val="1"/><c:axPos');
			expect(xml).toContain('sourceLinked="0"/><c:tickLblPos val="none"/><c:crossAx');
			expect(xml.match(/uri="kept"/g)).toHaveLength(2);
			expect(readChartFormatting(doc.documentElement)?.entries.categoryAxis).toMatchObject({
				axisVisible: false,
				labelsVisible: false,
			});
			expect(readChartFormatting(doc.documentElement)?.entries.valueAxis).toMatchObject({
				axisVisible: true,
				labelsVisible: false,
			});
		},
	);
	it('keeps labels hidden when revealing an imported hidden axis', () => {
		const doc = document(axis('catAx', 1, '<c:delete val="1"/><c:tickLblPos val="high"/>'));
		const formatting = readChartFormatting(doc.documentElement)!;
		formatting.entries.categoryAxis!.axisVisible = true;
		expect(writeChartAxisFormatting(doc.documentElement, formatting)).toBe(true);
		expect(readChartFormatting(doc.documentElement)?.entries.categoryAxis).toMatchObject({
			axisVisible: true,
			labelsVisible: false,
		});
		expect(buildXml(doc)).toContain('<c:tickLblPos val="none"/>');
	});

	it('keeps unspecified flags and does not manufacture axes for pie charts', () => {
		const doc = document(axis('catAx', 1, '<c:delete val="0"/><c:tickLblPos val="high"/>'));
		const before = buildXml(doc);
		expect(
			writeChartAxisFormatting(doc.documentElement, {
				sourceXml: '',
				entries: { categoryAxis: { sourceXml: '', labelsVisible: true } },
			}),
		).toBe(false);
		expect(buildXml(doc)).toBe(before);
		const pie = document('<c:pieChart/>');
		expect(
			writeChartAxisFormatting(pie.documentElement, {
				sourceXml: '',
				entries: { categoryAxis: { sourceXml: '', axisVisible: false } },
			}),
		).toBe(false);
		expect(buildXml(pie)).not.toContain('catAx');
	});
});
