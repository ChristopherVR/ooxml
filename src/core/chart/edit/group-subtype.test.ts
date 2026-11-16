import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { applyChartEditIntent } from './intent';
import { setChartGroupSubtype } from './group-subtype';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const part = (group: string) =>
	`<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea>${group}` +
	'<c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
	'<c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
	'</c:plotArea></c:chart></c:chartSpace>';
const ser = '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>';
const load = (group: string) => parseChartSpace(part(group)).chartSpace;
const reload = (space: ReturnType<typeof load>) =>
	parseChartSpace(writeChartSpace(space)).chartSpace.plotArea.groups[0]!;

describe('setChartGroupSubtype', () => {
	it('sets and removes the 3-D bar shape and gap depth', () => {
		const space = load(
			`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser}<c:gapDepth val="150"/><c:shape val="box"/><c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
		);
		const edited = setChartGroupSubtype(space, 0, { shape: 'cylinder', gapDepth: null });
		expect(reload(edited)).toMatchObject({ shape: 'cylinder' });
		expect(reload(edited).gapDepth).toBeUndefined();
		expect(() => setChartGroupSubtype(space, 0, { shape: 'star' as never })).toThrow(RangeError);
		expect(() => setChartGroupSubtype(space, 0, { gapDepth: 501 })).toThrow(RangeError);
		expect(() => setChartGroupSubtype(space, 0, { radarStyle: 'filled' })).toThrow(/no radarStyle/);
	});

	it('edits the radar style and surface wireframe in schema order', () => {
		const radar = load(
			`<c:radarChart><c:radarStyle val="marker"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:radarChart>`,
		);
		const xml = writeChartSpace(setChartGroupSubtype(radar, 0, { radarStyle: 'filled' }));
		expect(xml).toMatch(/<c:radarChart><c:radarStyle val="filled"\/><c:ser>/u);
		const surface = load(
			`<c:surfaceChart><c:wireframe val="0"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:surfaceChart>`,
		);
		expect(reload(setChartGroupSubtype(surface, 0, { wireframe: true })).wireframe).toBe(true);
	});

	it('writes a pie of pie split and its series lines', () => {
		const space = load(
			`<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>${ser}<c:gapWidth val="100"/><c:secondPieSize val="75"/></c:ofPieChart>`,
		);
		const edited = applyChartEditIntent(space, {
			op: 'groupSubtype',
			seriesIndex: 0,
			edit: {
				ofPieType: 'bar',
				splitType: 'cust',
				customSplit: [2, 3],
				secondPieSize: 50,
				seriesLines: true,
			},
		});
		const xml = writeChartSpace(edited);
		expect(xml).toContain('<c:ofPieType val="bar"/><c:varyColors val="1"/>');
		expect(xml).toMatch(
			/<c:gapWidth val="100"\/><c:splitType val="cust"\/><c:custSplit><c:secondPiePt val="2"\/><c:secondPiePt val="3"\/><\/c:custSplit><c:secondPieSize val="50"\/><c:serLines\/>/u,
		);
		expect(() => setChartGroupSubtype(space, 0, { secondPieSize: 4 })).toThrow(RangeError);
		expect(() => setChartGroupSubtype(space, 0, { customSplit: [-1] })).toThrow(RangeError);
		const removed = setChartGroupSubtype(edited, 0, { seriesLines: null, splitType: null });
		expect(writeChartSpace(removed)).not.toMatch(/serLines|splitType/u);
	});

	it('sets bubble options as unsigned numbers', () => {
		const space = load(
			`<c:bubbleChart><c:varyColors val="0"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:bubbleChart>`,
		);
		const xml = writeChartSpace(
			setChartGroupSubtype(space, 0, {
				bubbleScale: 150,
				showNegativeBubbles: false,
				sizeRepresents: 'w',
			}),
		);
		expect(xml).toContain(
			'<c:bubbleScale val="150"/><c:showNegBubbles val="0"/><c:sizeRepresents val="w"/>',
		);
		expect(() => setChartGroupSubtype(space, 0, { bubbleScale: 301 })).toThrow(RangeError);
		expect(() => setChartGroupSubtype(space, 0, { sizeRepresents: 'h' as never })).toThrow(
			RangeError,
		);
	});
});
