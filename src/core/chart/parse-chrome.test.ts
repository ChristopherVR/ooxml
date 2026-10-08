import { describe, expect, it } from 'vitest';
import { parseChartSpace } from './parse-space';
import { readDisplayNaAsBlank } from './parse-chrome';
import { writeChartSpace } from './write-space';

const ROOT =
	'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
const part = (chart: string) =>
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${ROOT}<c:chart>${chart}</c:chart></c:chartSpace>`;

const WALL =
	'<c:thickness val="0"/><c:spPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></c:spPr>';
const VALUE_AXIS =
	'<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/><c:crossBetween val="between"/><c:majorUnit val="1000"/>' +
	'<c:dispUnits><c:builtInUnit val="thousands"/><c:dispUnitsLbl><c:layout/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>k</a:t></a:r></a:p></c:rich></c:tx></c:dispUnitsLbl></c:dispUnits></c:valAx>';
const CATEGORY_AXIS =
	'<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>';
const CHART =
	'<c:autoTitleDeleted val="1"/>' +
	`<c:floor>${WALL}</c:floor><c:sideWall>${WALL}</c:sideWall><c:backWall><c:thickness val="5"/><c:pictureOptions><c:pictureFormat val="stretch"/></c:pictureOptions></c:backWall>` +
	'<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
	CATEGORY_AXIS +
	VALUE_AXIS +
	'<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="0"/><c:showOutline val="1"/><c:showKeys val="1"/><c:spPr><a:ln w="9525"/></c:spPr></c:dTable></c:plotArea>' +
	'<c:plotVisOnly val="1"/>';

describe('chart chrome: floor, walls, data table and display units', () => {
	it('reads them without reporting anything unmodelled', () => {
		const { chartSpace, issues } = parseChartSpace(part(CHART));
		expect(issues).toEqual([]);
		expect(chartSpace.floor).toMatchObject({ thickness: 0, spPr: { fill: { kind: 'solid' } } });
		expect(chartSpace.sideWall?.thickness).toBe(0);
		expect(chartSpace.backWall).toEqual({
			thickness: 5,
			pictureOptionsXml: expect.stringContaining('<c:pictureFormat val="stretch"/>'),
		});
		expect(chartSpace.plotArea.dataTable).toMatchObject({
			showHorizontalBorder: true,
			showVerticalBorder: false,
			showOutline: true,
			showKeys: true,
			spPr: { line: { widthEmu: 9525 } },
		});
		const units = chartSpace.plotArea.axes[1]?.displayUnits;
		expect(units?.builtInUnit).toBe('thousands');
		expect(units?.customUnit).toBeUndefined();
		expect(units?.label?.layout).toEqual({});
		expect(units?.label?.tx?.text).toBe('k');
	});

	it('reads a bare data-table switch as true (CT_Boolean default)', () => {
		const { chartSpace } = parseChartSpace(
			part('<c:plotArea><c:dTable><c:showKeys/></c:dTable></c:plotArea>'),
		);
		expect(chartSpace.plotArea.dataTable).toEqual({ showKeys: true });
	});

	it('writes them back byte for byte, in schema order', () => {
		const xml = part(CHART);
		expect(writeChartSpace(parseChartSpace(xml).chartSpace)).toBe(xml);
	});

	it('writes a custom unit instead of a built-in one', () => {
		const xml = part(
			`<c:plotArea>${CATEGORY_AXIS}<c:valAx><c:axId val="2"/><c:scaling/><c:crossAx val="1"/><c:dispUnits><c:custUnit val="250"/></c:dispUnits></c:valAx></c:plotArea>`,
		);
		const { chartSpace } = parseChartSpace(xml);
		expect(chartSpace.plotArea.axes[1]?.displayUnits).toEqual({ customUnit: 250 });
		expect(writeChartSpace(chartSpace)).toContain(
			'<c:dispUnits><c:custUnit val="250"/></c:dispUnits></c:valAx>',
		);
	});
});

describe('readDisplayNaAsBlank', () => {
	const ext = (flag: string) =>
		`<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16>${flag}</c16r3:dataDisplayOptions16></c:ext></c:extLst>`;

	it('reads the c16r3 flag from the chart extension list', () => {
		expect(readDisplayNaAsBlank(ext('<c16r3:dispNaAsBlank val="1"/>'))).toBe(true);
		expect(readDisplayNaAsBlank(ext('<c16r3:dispNaAsBlank val="0"/>'))).toBe(false);
		// CT_Boolean without @val means true.
		expect(readDisplayNaAsBlank(ext('<c16r3:dispNaAsBlank/>'))).toBe(true);
	});

	it('is undefined when the flag or the extension is not written', () => {
		expect(readDisplayNaAsBlank(undefined)).toBeUndefined();
		expect(readDisplayNaAsBlank(ext(''))).toBeUndefined();
		expect(
			readDisplayNaAsBlank(
				'<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ext uri="{X}"/></c:extLst>',
			),
		).toBeUndefined();
	});

	it('reads the flag from a parsed part', () => {
		const { chartSpace } = parseChartSpace(
			part(`<c:plotArea/>${ext('<c16r3:dispNaAsBlank val="1"/>')}`),
		);
		expect(readDisplayNaAsBlank(chartSpace.chartExtLst)).toBe(true);
	});
});
