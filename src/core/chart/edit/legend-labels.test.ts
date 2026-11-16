import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../parse-space';
import { setChartDataLabels } from './legend-labels';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const ser = (i: number) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/></c:ser>`;
const load = (labels: string) =>
	parseChartSpace(
		`<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/>` +
			`<c:grouping val="clustered"/>${ser(0)}${labels}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
			'<c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
			'<c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
			'</c:plotArea></c:chart></c:chartSpace>',
	).chartSpace;
const shown =
	'<c:dLbls><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
	'<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';

describe('setChartDataLabels on one series', () => {
	it('keeps what a series shows when only its number format changes', () => {
		const edited = setChartDataLabels(load(shown), { numberFormat: '0.0%' }, { series: 0 });
		expect(edited.plotArea.groups[0]!.series[0]!.dataLabels).toMatchObject({
			showValue: true,
			position: 'outEnd',
			numberFormat: { formatCode: '0.0%', sourceLinked: false },
		});
	});

	it('starts other edits from switched-off labels', () => {
		const edited = setChartDataLabels(load(shown), { showCategoryName: true }, { series: 0 });
		expect(edited.plotArea.groups[0]!.series[0]!.dataLabels).toMatchObject({
			showValue: false,
			showCategoryName: true,
		});
		const hidden = setChartDataLabels(load(''), { numberFormat: '0' }, { series: 0 });
		expect(hidden.plotArea.groups[0]!.series[0]!.dataLabels?.showValue).toBe(false);
	});
});
