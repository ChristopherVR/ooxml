// Generate a complete native/shared comparison manifest for the committed chart gallery.
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, posix, resolve } from 'node:path';
import JSZip from 'jszip';
import { parseXml, elements, NS } from '../../xml/index';
import { parseThemeXml } from '../../drawingml/theme';
import { parseChartSpace } from '../../chart/index';
import { readChartFormatting } from '../../chart/read-formatting';
import { readChartStyle } from '../../chart/read-style';
import { chartColorStyleId } from '../../chart/color-style';
import { chartExSummary } from '../../chart/render/summary-from-chartex';
import { chartSummaryFromSpace } from '../../chart/render/summary-from-space';
import { DRAWN_CHART_TYPES, renderChartSummarySvg } from '../../chart/render/render-space';

const directory = resolve(process.argv[2]!);
const source = resolve(process.argv[3] ?? 'pptx/__tests__/fixtures/e2e/chart-gallery.pptx');
mkdirSync(directory, { recursive: true });
const presentation = basename(source);
const target = resolve(directory, presentation);
if (target !== source) copyFileSync(source, target);
const zip = await JSZip.loadAsync(readFileSync(source));
const theme = parseThemeXml(await zip.file('ppt/theme/theme1.xml')!.async('string'));
const cases = [];
const skipped = [];
const slides = Object.keys(zip.files)
	.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
	.sort((a, b) => Number(a.match(/slide(\d+)/)![1]) - Number(b.match(/slide(\d+)/)![1]));
for (const part of slides) {
	const slide = Number(part.match(/slide(\d+)/)![1]);
	const document = parseXml(await zip.file(part)!.async('string'));
	const relsPath = posix.join(posix.dirname(part), '_rels', `${posix.basename(part)}.rels`);
	const rels = parseXml(await zip.file(relsPath)!.async('string'));
	const tree = document.getElementsByTagNameNS(NS.p, 'spTree')[0]!;
	let shape = 0;
	for (const element of elements(tree)) {
		if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(element.localName)) continue;
		shape++;
		const chart = [...element.getElementsByTagName('*')].find((node) => node.localName === 'chart');
		if (!chart) continue;
		const id = chart.getAttributeNS(NS.r, 'id');
		const relationship = elements(rels.documentElement).find(
			(node) => node.getAttribute('Id') === id,
		)!;
		const chartPart = posix.normalize(
			posix.join('ppt/slides', relationship.getAttribute('Target')!),
		);
		const xml = await zip.file(chartPart)!.async('string');
		const chartDocument = parseXml(xml);
		const extended = chartDocument.documentElement.namespaceURI?.includes('/2014/chartex');
		const extras: import('../../chart/render/summary-from-space').ChartSummaryExtras = {};
		const chartRels = zip.file(
			posix.join(posix.dirname(chartPart), '_rels', `${posix.basename(chartPart)}.rels`),
		);
		if (chartRels)
			for (const rel of elements(parseXml(await chartRels.async('string')).documentElement)) {
				if (rel.getAttribute('TargetMode') === 'External') continue;
				const linked = zip.file(
					posix.normalize(posix.join(posix.dirname(chartPart), rel.getAttribute('Target')!)),
				);
				if (!linked) continue;
				const type = rel.getAttribute('Type') ?? '';
				if (type.endsWith('/chartStyle'))
					extras.styleDefinition = readChartStyle(await linked.async('string'));
				if (type.endsWith('/chartColorStyle'))
					extras.colorPalette = chartColorStyleId(await linked.async('string'));
			}
		extras.formatting = readChartFormatting(chartDocument.documentElement);
		const summary = extended
			? chartExSummary(xml)?.summary
			: chartSummaryFromSpace(parseChartSpace(xml).chartSpace, extras);
		if (!summary || !DRAWN_CHART_TYPES.has(summary.chartType)) {
			skipped.push({ slide, shape, chartPart, reason: 'No shared painter for this chart family' });
			continue;
		}
		const caseId = `chart-${slide}-${shape}-${summary.chartType.toLowerCase()}`;
		const width = 840;
		const height = 420;
		const sharedSvg = `${caseId}.svg`;
		writeFileSync(
			resolve(directory, sharedSvg),
			renderChartSummarySvg(summary, {
				width,
				height,
				theme: theme.colorScheme,
				fonts: { major: theme.fontScheme.major.latin, minor: theme.fontScheme.minor.latin },
			}),
		);
		cases.push({ id: caseId, presentation, slide, shape, width, height, sharedSvg, chartPart });
	}
}
writeFileSync(
	resolve(directory, 'manifest.json'),
	JSON.stringify(
		{
			profile: 'Shared chart summary painter with document theme, native full chart comparison',
			cases,
			skipped,
		},
		null,
		2,
	),
);
console.log(`${cases.length} chart cases; ${skipped.length} explicitly skipped`);
