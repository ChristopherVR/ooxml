import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { chartSummaryView } from './chart-view';
import { chartThemePalette, renderChartSpaceSvg } from './render-space';
import { chartSummaryFromSpace } from './summary-from-space';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const num = (values: number[]) =>
	`<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}`;
const cats = (labels: string[]) =>
	`<c:cat><c:strRef><c:f>S!$A$2:$A$9</c:f><c:strCache><c:ptCount val="${labels.length}"/>${labels.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;
const series = (
	idx: number,
	name: string,
	values: number[],
	labels = '',
	category = ['A', 'B', 'C'],
) =>
	`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:v>${name}</c:v></c:tx>${labels}${cats(category)}<c:val><c:numRef><c:f>S!$B$2:$B$9</c:f><c:numCache><c:formatCode>General</c:formatCode>${num(values)}</c:numCache></c:numRef></c:val></c:ser>`;
const AXES = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`;
const flags = (on: string[], extra = '') =>
	['showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']
		.map((flag) => `<c:${flag} val="${on.includes(flag) ? 1 : 0}"/>`)
		.join('') + extra;
const dLbls = (body: string) => `<c:dLbls>${body}</c:dLbls>`;

const chart = (group: string, withAxes = true) =>
	`<c:chartSpace ${C}><c:chart><c:plotArea><c:layout/>${group}${withAxes ? AXES : ''}</c:plotArea></c:chart></c:chartSpace>`;
const column = (body: string, grouping = 'clustered', dir = 'col') =>
	chart(
		`<c:barChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/>${body}<c:axId val="1"/><c:axId val="2"/></c:barChart>`,
	);
const render = (xml: string) =>
	renderChartSpaceSvg(parseChartSpace(xml).chartSpace, { width: 400, height: 300 });
const texts = (svg: string) =>
	[...svg.matchAll(/<text [^>]*>([^<]*)<\/text>/gu)].map((m) => m[1] ?? '');
/** The `y` of the first text element whose content is `content`. */
const textY = (svg: string, content: string) =>
	Number(
		new RegExp(`<text x="[\\d.-]+" y="([\\d.-]+)"[^>]*>${content}</text>`, 'u').exec(svg)?.[1],
	);
const textX = (svg: string, content: string) =>
	Number(
		new RegExp(`<text x="([\\d.-]+)" y="[\\d.-]+"[^>]*>${content}</text>`, 'u').exec(svg)?.[1],
	);

describe('data labels in the summary and view', () => {
	it('carries the group and series levels and prunes an all-zero group', () => {
		const zero = parseChartSpace(column(series(0, 'S', [1, 2, 3]) + dLbls(flags([])))).chartSpace;
		expect(chartSummaryFromSpace(zero).series[0]?.dataLabels).toBeUndefined();
		const own = parseChartSpace(
			column(series(0, 'S', [1, 2, 3], dLbls(flags(['showVal'], '<c:dLblPos val="inEnd"/>')))),
		).chartSpace;
		expect(chartSummaryFromSpace(own).series[0]?.dataLabels?.series).toMatchObject({
			showValue: true,
			position: 'inEnd',
		});
	});

	it('builds the text of every point through the cascade', () => {
		const space = parseChartSpace(
			column(
				series(
					0,
					'Sales',
					[10, 20, 30],
					dLbls(
						'<c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl>' +
							'<c:numFmt formatCode="0.0&quot; kg&quot;" sourceLinked="0"/>' +
							flags(['showVal', 'showCatName']),
					),
				),
			),
		).chartSpace;
		const view = chartSummaryView(chartSummaryFromSpace(space), chartThemePalette(undefined));
		expect(view.series[0]?.labels?.map((l) => l?.text)).toEqual([
			'A, 10.0 kg',
			undefined,
			'C, 30.0 kg',
		]);
	});

	it('shows percentages of the series total with a line break after the category', () => {
		const space = parseChartSpace(
			chart(
				`<c:pieChart>${series(0, 'S', [1, 3], dLbls(flags(['showCatName', 'showPercent'])), ['x', 'y'])}</c:pieChart>`,
				false,
			),
		).chartSpace;
		const view = chartSummaryView(chartSummaryFromSpace(space), chartThemePalette(undefined));
		expect(view.series[0]?.labels?.map((l) => l?.text)).toEqual(['x\n25%', 'y\n75%']);
	});
});

describe('data labels in the SVG', () => {
	it('draws no label without c:dLbls, and the values with them', () => {
		const plain = render(column(series(0, 'S', [10, 20, 30])));
		const labelled = render(column(series(0, 'S', [10, 20, 30], dLbls(flags(['showVal'])))));
		expect(texts(labelled).length - texts(plain).length).toBe(3);
		expect(texts(labelled)).toEqual(expect.arrayContaining(['10', '20', '30']));
	});

	it('places column labels above (outEnd), inside the end, inside the base and centred', () => {
		const at = (position: string) =>
			textY(
				render(
					column(
						series(
							0,
							'S',
							[11, 22, 33.3],
							dLbls(flags(['showVal'], `<c:dLblPos val="${position}"/>`)),
						),
					),
				),
				'33.3',
			);
		const [outEnd, inEnd, ctr, inBase] = ['outEnd', 'inEnd', 'ctr', 'inBase'].map(at) as [
			number,
			number,
			number,
			number,
		];
		expect(outEnd).toBeLessThan(inEnd);
		expect(inEnd).toBeLessThan(ctr);
		expect(ctr).toBeLessThan(inBase);
	});

	it('defaults to the centre of a stacked segment and to outside the end of clustered bars', () => {
		const body = (grouping: string) =>
			column(
				series(0, 'A', [10.5, 10.5, 10.5], dLbls(flags(['showVal']))) +
					series(1, 'B', [33.3, 33.3, 33.3], dLbls(flags(['showVal']))),
				grouping,
			);
		const stacked = render(body('stacked'));
		const clustered = render(body('clustered'));
		// Stacked: series B (30) is centred inside its segment, below the plot top.
		expect(textY(stacked, '33.3')).toBeGreaterThan(40);
		expect(texts(clustered).filter((t) => t === '33.3')).toHaveLength(3);
		// Outside the end of a clustered bar is above the centre of the same bar when stacked.
		expect(textY(clustered, '33.3')).toBeLessThan(textY(stacked, '33.3'));
	});

	it('puts a horizontal bar label past the bar end', () => {
		const svg = render(
			column(series(0, 'S', [11, 22, 33.3], dLbls(flags(['showVal']))), 'clustered', 'bar'),
		);
		const bar = /data-chart-point="2"><rect x="([\d.]+)" y="[\d.]+" width="([\d.]+)"/u.exec(svg);
		const rect = Number(bar?.[1]) + Number(bar?.[2]);
		expect(textX(svg, '33.3')).toBeGreaterThan(rect - 1);
	});

	it('draws pie labels: the percentage inside a big slice and best fit outside a thin one', () => {
		const svg = render(
			chart(
				`<c:pieChart>${series(0, 'S', [97, 1, 2], dLbls(flags(['showPercent'])), ['a', 'b', 'c'])}</c:pieChart>`,
				false,
			),
		);
		expect(texts(svg)).toEqual(expect.arrayContaining(['97%', '1%', '2%']));
	});

	it('labels a line chart at its points, to the right by default', () => {
		const svg = render(
			chart(
				`<c:lineChart><c:grouping val="standard"/>${series(0, 'S', [5.25, 8.5, 3.5], dLbls(flags(['showVal'])))}<c:axId val="1"/><c:axId val="2"/></c:lineChart>`,
			),
		);
		const circle = /<circle cx="([\d.]+)" cy="([\d.]+)" r="2.5"/u.exec(svg);
		expect(textX(svg, '5.25')).toBeGreaterThan(Number(circle?.[1]));
	});

	it('escapes label text', () => {
		const svg = render(
			column(
				series(
					0,
					'S',
					[1],
					dLbls(
						'<c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>&lt;b&gt; &amp; co</a:t></a:r></a:p></c:rich></c:tx>' +
							flags(['showVal']) +
							'</c:dLbl>' +
							flags(['showVal']),
					),
					['A'],
				),
			),
		);
		expect(svg).toContain('&lt;b&gt; &amp; co');
	});
});
