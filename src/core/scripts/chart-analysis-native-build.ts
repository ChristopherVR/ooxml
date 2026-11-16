// Native acceptance files for the shared chart analysis and picture fill edits (U2), one per host:
// a clustered column chart with two trendlines and custom error values on its first series, a
// picture fill on its second series and a tiled picture on its plot area, and a stacked column
// chart whose last series moved to the secondary axis. Word and Excel get the edits through
// `applyChartEditIntent` (what their Format panes send), PowerPoint through its own chart
// operations (what its pane applies: withChartSeriesTrendline, withChartSeriesErrorBars). Each file is reopened by its own host first; manifest.json
// lists what `chart-analysis-native-acceptance.ps1` checks through each application's COM model.
//
//   bun scripts/chart-analysis-native-build.ts <output folder>
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';
import { applyChartEditIntent, type ChartEditIntent } from '../chart/edit/intent';
import { chartErrorBarSourceFromText } from '../chart/edit/error-bar-source';
import { chartSpacePictureIds } from '../chart/picture-parts';
import {
	chartsIn,
	createDocument,
	documentDrawingBlocks,
	editDocxChart,
	loadDocx,
	newChartImage,
	newDocxChart,
	saveDocx,
	type DocxChart,
} from '../docx/index';
import type { TextRun } from '../docx/model';
import { createEditSession } from '../xlsx/edit/session';
import type { ChartObject } from '../xlsx/model';
import { loadXlsx } from '../xlsx/read/index';
import { createWorkbook } from '../xlsx/workbook';
import { saveXlsx } from '../xlsx/write/index';
import { PptxHandler } from '../pptx/core/PptxHandler';
import { withChartElementFormat } from '../pptx/core/builders/sdk/chart-element-format-operations';
import { withChartSeriesErrorBars } from '../pptx/core/builders/sdk/chart-error-bar-operations';
import { withChartSeriesTrendline } from '../pptx/core/builders/sdk/chart-trendline-operations';
import { withChartSeriesOnSecondaryAxis } from '../pptx/core/builders/sdk/chart-series-axis-operations';
import type { ChartPptxElement, PptxChartData } from '../pptx/core/types';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/chart-analysis-native-build.ts <output folder>');
mkdirSync(out, { recursive: true });

/** A 16 x 16 two-colour PNG, as a `data:` URL. */
function picture(): string {
	const crc = (bytes: Buffer) => {
		let c = ~0;
		for (const byte of bytes) {
			c ^= byte;
			for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
		}
		return ~c >>> 0;
	};
	const chunk = (type: string, data: Buffer) => {
		const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
		const head = Buffer.alloc(4);
		head.writeUInt32BE(data.length);
		const tail = Buffer.alloc(4);
		tail.writeUInt32BE(crc(body));
		return Buffer.concat([head, body, tail]);
	};
	const header = Buffer.alloc(13);
	header.writeUInt32BE(16, 0);
	header.writeUInt32BE(16, 4);
	header.set([8, 2, 0, 0, 0], 8);
	const rows = Array.from({ length: 16 }, (_, y) =>
		Buffer.from([
			0,
			...Array.from({ length: 16 }, (_, x) =>
				((x >> 2) + (y >> 2)) % 2 ? [0xc0, 0x50, 0x4d] : [0xf5, 0xd9, 0x8b],
			).flat(),
		]),
	);
	const png = Buffer.concat([
		Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
		chunk('IHDR', header),
		chunk('IDAT', deflateSync(Buffer.concat(rows))),
		chunk('IEND', Buffer.alloc(0)),
	]);
	return `data:image/png;base64,${png.toString('base64')}`;
}

const PICTURE = picture();
const literal = (text: string, pointCount: number) =>
	chartErrorBarSourceFromText(text, { pointCount })!;

/** The analysis and picture edits, for a chart of `points` categories. */
const analysis = (points: number): ChartEditIntent[] => [
	{ op: 'trendline', seriesIndex: 0, index: 0, edit: { type: 'linear' } },
	{ op: 'trendline', seriesIndex: 0, index: 1, edit: { type: 'movingAvg', period: 2 } },
	{
		op: 'errorBars',
		seriesIndex: 0,
		direction: 'y',
		edit: {
			valueType: 'cust',
			plus: literal(Array.from({ length: points }, (_, i) => i + 1).join(', '), points),
			minus: literal('0.5', points),
		},
	},
	{ op: 'picture', target: { kind: 'series', seriesIndex: 1 }, edit: { src: PICTURE } },
	{ op: 'picture', target: { kind: 'plotArea' }, edit: { src: PICTURE, tile: true } },
];
const STACKED: ChartEditIntent[] = [{ op: 'secondaryAxis', seriesIndex: 2, secondary: true }];

/** What a chart must show in the application, by 1-based series number. */
interface NativeChart {
	series: number;
	/** Series on the secondary axis (`AxisGroup` 2); every other series is primary. */
	secondary: number[];
	trendlines?: { series: number; count: number };
	errorBars?: { series: number };
	/** Series with a picture fill, and whether the plot area has one. */
	pictureSeries?: number[];
	pictureArea?: boolean;
}
const ANALYSED: NativeChart = {
	series: 3,
	secondary: [],
	trendlines: { series: 1, count: 2 },
	errorBars: { series: 1 },
	pictureSeries: [2],
	pictureArea: true,
};
const STACKED_SECONDARY: NativeChart = { series: 3, secondary: [3] };
const manifest: Record<
	string,
	{ host: 'word' | 'excel' | 'powerpoint'; slide?: number; charts: NativeChart[] }
> = {};

// Word: two new charts in a blank document, saved by `saveDocx` and reopened.
{
	const doc = createDocument();
	const add = (chart: DocxChart) =>
		doc.blocks.push({
			type: 'paragraph',
			id: `chart-${doc.blocks.length}`,
			runs: [{ text: '', image: newChartImage(chart) } as TextRun],
		});
	const column = newDocxChart(undefined, { title: 'Analysis and pictures' });
	const points = column.chartSpace!.plotArea.groups[0]!.series[0]!.values!.cache!.points.length;
	add(editDocxChart(column, (space) => analysis(points).reduce(applyChartEditIntent, space)));
	add(
		editDocxChart(
			newDocxChart({ kind: 'bar', grouping: 'stacked' }, { title: 'Stacked' }),
			(space) => STACKED.reduce(applyChartEditIntent, space),
		),
	);
	const bytes = await saveDocx(doc);
	const reopened = chartsIn(documentDrawingBlocks((await loadDocx(bytes)).model));
	if (reopened.length !== 2 || !Object.keys(reopened[0]!.chartSpace?.pictureSources ?? {}).length)
		throw new Error('Word: own-host reopen lost a chart or its pictures');
	writeFileSync(join(out, 'word-analysis.docx'), bytes);
	manifest['word-analysis.docx'] = { host: 'word', charts: [ANALYSED, STACKED_SECONDARY] };
}

// Excel: two charts over one data block, edited through the edit session and reopened.
{
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['', 'Series 1', 'Series 2', 'Series 3'],
		['Category 1', 4.3, 2.4, 2],
		['Category 2', 2.5, 4.4, 2],
		['Category 3', 3.5, 1.8, 3],
		['Category 4', 4.5, 2.8, 5],
	]);
	const series = [1, 2, 3].map((n) => {
		const col = 'BCD'[n - 1]!;
		return {
			name: `Series ${n}`,
			nameRef: `Sheet1!$${col}$1`,
			valuesRef: `Sheet1!$${col}$2:$${col}$5`,
			categoriesRef: 'Sheet1!$A$2:$A$5',
			values: [],
			categories: [],
		};
	});
	for (const [row, grouping] of [
		[0, 'clustered'],
		[20, 'stacked'],
	] as const)
		session.addChart(0, {
			anchor: { from: { row, col: 6, rowOffset: 0, colOffset: 0 }, ext: { cx: 5e6, cy: 3e6 } },
			chartType: 'column',
			grouping,
			showLegend: true,
			series: structuredClone(series),
		});
	for (const intent of analysis(4)) session.editChart(0, 0, intent);
	for (const intent of STACKED) session.editChart(0, 1, intent);
	const bytes = await saveXlsx(workbook);
	const charts = (await loadXlsx(bytes)).sheets[0]!.drawings.filter(
		(d): d is ChartObject => d.kind === 'chart',
	);
	const ids = charts[0]?.chartSpace ? chartSpacePictureIds(charts[0].chartSpace) : [];
	if (charts.length !== 2 || ids.length === 0)
		throw new Error('Excel: own-host reopen lost a chart or its pictures');
	writeFileSync(join(out, 'excel-analysis.xlsx'), bytes);
	manifest['excel-analysis.xlsx'] = { host: 'excel', charts: [ANALYSED, STACKED_SECONDARY] };
}

// PowerPoint: the column chart of a PowerPoint-authored fixture, edited through the operations its
// Format pane applies and saved by the handler, once with the analysis and once stacked.
{
	const source = new Uint8Array(
		readFileSync(
			new URL('../pptx/__tests__/fixtures/e2e/chart-data-fidelity.pptx', import.meta.url),
		),
	);
	const deckFor = async () => {
		const handler = new PptxHandler();
		const deck = await handler.load(source);
		const slide = deck.slides.findIndex((item) =>
			item.elements.some((entry) => entry.type === 'chart' && entry.chartData?.chartType === 'bar'),
		);
		const element = deck.slides[slide]?.elements.find((item) => item.type === 'chart') as
			| ChartPptxElement
			| undefined;
		if (!element?.chartData) throw new Error('No native column chart found.');
		element.chartData.categories = ['Plan', 'Design', 'Build', 'Review', 'Ship'];
		element.chartData.series = [
			{ name: 'Series 1', values: [2, 4, 6, 8, 10] },
			{ name: 'Series 2', values: [3, 6, 9, 12, 15] },
			{ name: 'Series 3', values: [1, 5, 7, 11, 13] },
		];
		return { handler, deck, element, slide };
	};
	const save = async (
		name: string,
		edit: (data: PptxChartData) => PptxChartData,
		want: NativeChart,
	) => {
		const { handler, deck, element, slide } = await deckFor();
		element.chartData = edit(element.chartData!);
		const bytes = await handler.save(deck.slides);
		const back = (await new PptxHandler().load(bytes)).slides[slide]!.elements.find(
			(item) => item.type === 'chart',
		) as ChartPptxElement;
		if (back.chartData?.series.length !== 3) throw new Error(`${name}: own-host reopen differs`);
		writeFileSync(join(out, name), bytes);
		manifest[name] = { host: 'powerpoint', slide: slide + 1, charts: [want] };
	};
	await save(
		'powerpoint-analysis.pptx',
		(data) => {
			// The trendline intents the shared pane sends, through PowerPoint's operation.
			let next = withChartSeriesTrendline(data, 0, 0, { type: 'linear' });
			next = withChartSeriesTrendline(next, 0, 1, { type: 'movingAvg', period: 2 });
			next = withChartSeriesErrorBars(next, 0, {
				valueType: 'cust',
				plus: literal('1, 2, 3, 4, 5', 5),
				minus: literal('0.5', 5),
			});
			next = withChartElementFormat(
				next,
				{ kind: 'series', seriesIndex: 1 },
				{
					fill: { kind: 'picture', picture: { src: PICTURE } },
				},
			);
			return withChartElementFormat(
				next,
				{ kind: 'plotArea' },
				{
					fill: { kind: 'picture', picture: { src: PICTURE, tile: true } },
				},
			);
		},
		ANALYSED,
	);
	await save(
		'powerpoint-stacked-secondary.pptx',
		(data) => withChartSeriesOnSecondaryAxis({ ...data, grouping: 'stacked' }, 2, true),
		STACKED_SECONDARY,
	);
}

writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`Wrote ${Object.keys(manifest).length} files to ${out}`);
