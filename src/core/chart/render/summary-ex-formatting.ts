import type { ChartExSpace } from '../ex/model';
import type { ChartTextBody } from '../model-series';
import type { ChartStyleEntry } from '../style-definition';
import type { ChartSummary } from './summary';
import { chartRunStyle } from './chart-title-text';
function textStyle(body: ChartTextBody | undefined): Partial<ChartStyleEntry> {
	const paragraph = body?.paragraphs[0];
	return {
		...chartRunStyle(paragraph?.defaultProperties ?? {}),
		...chartRunStyle(paragraph?.runs[0] ?? {}),
	};
}
/** Office ChartEx defaults, overridden by authored DrawingML text formatting. */
export function chartExFormatting(space: ChartExSpace): NonNullable<ChartSummary['formatting']> {
	return {
		builtInStyle: 2,
		sourceXml: '',
		entries: {
			title: {
				fontSize: 18,
				bold: false,
				textColor: {
					kind: 'scheme',
					value: 'tx1',
					transforms: [
						{ name: 'lumMod', value: '65000' },
						{ name: 'lumOff', value: '35000' },
					],
				},
				...textStyle(space.title?.txPr),
				...textStyle(space.title?.tx?.rich),
				...(space.title?.tx?.rich ? { textBody: space.title.tx.rich } : {}),
				sourceXml: '',
			},
			legend: { fontSize: 12, ...textStyle(space.legend?.txPr), sourceXml: '' },
			categoryAxis: {
				fontSize: 10,
				...textStyle(space.plotArea.axes.find((a) => a.scaling?.kind === 'cat')?.txPr),
				sourceXml: '',
			},
		},
	};
}
