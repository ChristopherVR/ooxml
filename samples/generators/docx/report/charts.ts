/**
 * The report's two native charts, built with the core's chart model: a new chart from a data
 * sheet (`newDocxChart`, which also writes the embedded workbook on save), then styled through the
 * model's shape and text properties.
 */
import type { TextRun } from '../../../../src/core/docx/index';
import type {
	ChartAxis,
	ChartShapeProperties,
	ChartSpace,
	ChartTextBody,
} from '../../../../src/core/chart/index';
import type { DrawingColor } from '../../../../src/core/drawingml/types';
import { FONTS } from '../../lib/fiction';
import { chart, pxFromCm } from '../kit/drawing';
import { MONTHS, PALETTE, POWER_CURVE } from './data';

const rgb = (value: string): DrawingColor => ({ kind: 'srgb', value, transforms: [] });
const fill = (value: string): ChartShapeProperties => ({
	fill: { kind: 'solid', color: rgb(value) },
	line: { fill: { kind: 'none' } },
});
const stroke = (value: string, widthPt: number): ChartShapeProperties => ({
	fill: { kind: 'none' },
	line: { widthEmu: Math.round(widthPt * 12700), fill: { kind: 'solid', color: rgb(value) } },
});

/** Text properties: one empty paragraph that carries the default run formatting. */
const text = (sizePt: number, color: string, bold = false): ChartTextBody => ({
	paragraphs: [
		{ defaultProperties: { sizePt, bold, color: rgb(color), typeface: FONTS.body }, runs: [] },
	],
	text: '',
});

function axisTitle(label: string): NonNullable<ChartAxis['title']> {
	return {
		tx: {
			rich: {
				paragraphs: [
					{
						defaultProperties: {
							sizePt: 8.5,
							bold: false,
							color: rgb(PALETTE.slate),
							typeface: FONTS.body,
						},
						runs: [{ text: label }],
					},
				],
				text: label,
			},
			text: label,
		},
		text: label,
		overlay: false,
	};
}

/** The look both charts share: no frame, quiet gridlines, the report's sans for every label. */
function frame(draft: ChartSpace, titles: { category: string; value: string }): void {
	draft.roundedCorners = false;
	draft.spPr = { fill: { kind: 'none' }, line: { fill: { kind: 'none' } } };
	draft.txPr = text(8.5, PALETTE.slate);
	draft.plotArea.spPr = { fill: { kind: 'none' }, line: { fill: { kind: 'none' } } };
	for (const axis of draft.plotArea.axes) {
		axis.txPr = text(8.5, PALETTE.slate);
		axis.majorTickMark = 'none';
		axis.minorTickMark = 'none';
		if (axis.kind === 'val') {
			axis.majorGridlines = true;
			axis.majorGridlinesSpPr = stroke(PALETTE.rule, 0.5);
			axis.spPr = { fill: { kind: 'none' }, line: { fill: { kind: 'none' } } };
			axis.title = axisTitle(titles.value);
			axis.numberFormat = { formatCode: '#,##0', sourceLinked: false };
		} else {
			axis.majorGridlines = false;
			axis.spPr = stroke(PALETTE.slate, 0.75);
			axis.title = axisTitle(titles.category);
		}
	}
	if (draft.legend) {
		draft.legend.position = 't';
		draft.legend.txPr = text(8.5, PALETTE.ink);
	}
}

const SIZE = { widthPx: pxFromCm(16), heightPx: pxFromCm(8.2) };

/** Figure 2: modelled against measured net energy, month by month (clustered columns). */
export function energyChart(): TextRun {
	return chart(
		{
			categories: MONTHS.map((month) => month.label),
			series: [
				{ name: 'Modelled', values: MONTHS.map((month) => month.modelled) },
				{ name: 'Measured', values: MONTHS.map((month) => month.measured) },
			],
		},
		{ kind: 'bar', barDirection: 'col', grouping: 'clustered' },
		{ legendPosition: 't', group: { gapWidth: 70, overlap: -10 } },
		{
			...SIZE,
			name: 'Monthly net energy chart',
			alt: 'Clustered column chart of modelled and measured net energy for each month of the first year.',
		},
		(draft) => {
			frame(draft, { category: 'Month, November 2024 to October 2025', value: 'Net energy (MWh)' });
			const [modelled, measured] = draft.plotArea.groups[0]!.series;
			modelled!.spPr = fill(PALETTE.sea);
			measured!.spPr = fill(PALETTE.deep);
			const values = draft.plotArea.axes.find((axis) => axis.kind === 'val');
			if (values) {
				values.scaling = { ...values.scaling, min: 0, max: 600 };
				values.majorUnit = 100;
			}
		},
	);
}

/** Figure 3: the binned power curve of one turbine against the model (lines with markers). */
export function powerCurveChart(): TextRun {
	return chart(
		{
			categories: POWER_CURVE.map((bin) => bin.speed.toFixed(1)),
			series: [
				{ name: 'Modelled', values: POWER_CURVE.map((bin) => bin.modelled) },
				{ name: 'Measured (T1 to T4 mean)', values: POWER_CURVE.map((bin) => bin.measured) },
			],
		},
		{ kind: 'line', grouping: 'standard' },
		{ legendPosition: 't' },
		{
			...SIZE,
			name: 'Power curve chart',
			alt: 'Line chart of modelled and measured turbine power against flow speed, reaching 500 kW near 2.3 metres per second.',
		},
		(draft) => {
			frame(draft, { category: 'Flow speed at hub height (m/s)', value: 'Electrical power (kW)' });
			const group = draft.plotArea.groups[0]!;
			group.marker = true;
			const [modelled, measured] = group.series;
			modelled!.spPr = stroke(PALETTE.sea, 2.25);
			modelled!.marker = { symbol: 'none' };
			modelled!.smooth = false;
			measured!.spPr = stroke(PALETTE.amber, 1.5);
			measured!.marker = {
				symbol: 'circle',
				size: 6,
				spPr: {
					fill: { kind: 'solid', color: rgb(PALETTE.amber) },
					line: { widthEmu: 9525, fill: { kind: 'solid', color: rgb(PALETTE.white) } },
				},
			};
			measured!.smooth = false;
			const values = draft.plotArea.axes.find((axis) => axis.kind === 'val');
			if (values) {
				values.scaling = { ...values.scaling, min: 0, max: 600 };
				values.majorUnit = 100;
			}
		},
	);
}
