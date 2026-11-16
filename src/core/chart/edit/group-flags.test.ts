import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { setChartBarDirection, setChartVaryColors } from './group-flags';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const ser = (i: number) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/></c:ser>`;
const axes =
	'<c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
	'<c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>';
const load = (groups: string) =>
	parseChartSpace(
		`<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea>${groups}${axes}</c:plotArea></c:chart></c:chartSpace>`,
	).chartSpace;
const bar = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser(0)}<c:axId val="1"/><c:axId val="2"/></c:barChart>`;
const line = `<c:lineChart><c:grouping val="standard"/>${ser(1)}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`;

describe('setChartVaryColors', () => {
	it('sets and removes the flag in schema order', () => {
		const xml = writeChartSpace(setChartVaryColors(load(bar), 0, true));
		expect(xml).toMatch(/<c:grouping val="clustered"\/><c:varyColors val="1"\/><c:ser>/u);
		expect(writeChartSpace(setChartVaryColors(load(bar), 0, null))).not.toContain('varyColors');
	});

	it('refuses a group without the setting', () => {
		const stock = load(
			`<c:stockChart>${ser(0)}${ser(1)}${ser(2)}<c:axId val="1"/><c:axId val="2"/></c:stockChart>`,
		);
		expect(() => setChartVaryColors(stock, 0, true)).toThrow(/no varyColors/u);
	});
});

describe('setChartBarDirection', () => {
	it('turns columns into bars and swaps the axis sides', () => {
		const edited = setChartBarDirection(load(bar), 'bar');
		expect(edited.plotArea.groups[0]!.barDirection).toBe('bar');
		expect(edited.plotArea.axes.map((axis) => axis.position)).toEqual(['l', 'b']);
		const back = setChartBarDirection(edited, 'col');
		expect(back.plotArea.axes.map((axis) => axis.position)).toEqual(['b', 'l']);
	});

	it('refuses a combination with another family', () => {
		expect(() => setChartBarDirection(load(bar + line), 'bar')).toThrow(/bar groups/u);
		expect(() => setChartBarDirection(load(bar), 'row' as never)).toThrow(RangeError);
	});
});
