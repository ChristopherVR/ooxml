// Round trip of the model gaps the chart edit operations needed (docs/agnostic-core-plan.md, "Model
// gaps"): pie-of-pie split options and date-axis time units, written in the schema order.
import { describe, expect, it } from 'vitest';
import { NS } from '../xml/index';
import { parseChartSpace } from './parse-space';
import { writeChartSpace } from './write-space';

const part = (plot: string) =>
	`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:chart><c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`;

describe('chart model gaps', () => {
	it('reads and writes the pie-of-pie split options', () => {
		const xml = part(
			'<c:ofPieChart><c:ofPieType val="bar"/><c:varyColors val="1"/><c:gapWidth val="100"/>' +
				'<c:splitType val="cust"/><c:splitPos val="2"/><c:custSplit><c:secondPiePt val="3"/>' +
				'<c:secondPiePt val="4"/></c:custSplit><c:secondPieSize val="60"/><c:serLines/></c:ofPieChart>',
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(issues).toEqual([]);
		const group = chartSpace.plotArea.groups[0]!;
		expect(group).toMatchObject({
			splitType: 'cust',
			splitPosition: 2,
			customSplit: [3, 4],
			secondPieSize: 60,
		});
		const written = writeChartSpace(chartSpace);
		expect(written).toContain(
			'<c:gapWidth val="100"/><c:splitType val="cust"/><c:splitPos val="2"/><c:custSplit>' +
				'<c:secondPiePt val="3"/><c:secondPiePt val="4"/></c:custSplit><c:secondPieSize val="60"/>' +
				'<c:serLines/>',
		);
		expect(parseChartSpace(written).chartSpace).toEqual(chartSpace);
	});

	it('reads and writes the date axis time units', () => {
		const xml = part(
			'<c:dateAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
				'<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/><c:auto val="1"/>' +
				'<c:lblOffset val="100"/><c:baseTimeUnit val="days"/><c:majorUnit val="1"/>' +
				'<c:majorTimeUnit val="months"/><c:minorUnit val="7"/><c:minorTimeUnit val="days"/>' +
				'</c:dateAx>',
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(issues).toEqual([]);
		expect(chartSpace.plotArea.axes[0]).toMatchObject({
			majorTimeUnit: 'months',
			minorTimeUnit: 'days',
		});
		const written = writeChartSpace(chartSpace);
		expect(written).toContain(
			'<c:majorUnit val="1"/><c:majorTimeUnit val="months"/><c:minorUnit val="7"/>' +
				'<c:minorTimeUnit val="days"/>',
		);
		expect(parseChartSpace(written).chartSpace).toEqual(chartSpace);
	});
});
