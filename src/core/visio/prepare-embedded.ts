import type JSZip from 'jszip';
import { openZip } from '../opc/zip-limits';
import {
	chartExSummaryFromModel,
	chartSummaryView,
	chartThemePalette,
	parseChartExSpace,
	type ChartViewModel,
} from '../chart/index';
import { parseThemeXml } from '../drawingml/theme';
import { NS } from '../xml/index';
import { VISIO_CHART_OBJECT_MAX_BYTES } from './edit-chart-object-commands';
import type { VisioEmbeddedObject } from './model';
import type { VisioPackage } from './package';
import { children, type Report } from './sheet';

const CHART_EX_VIEW_TYPES = new Set<ChartViewModel['type']>([
	'waterfall',
	'histogram',
	'treemap',
	'sunburst',
	'funnel',
	'boxWhisker',
	'regionMap',
]);

const CHART_EX_PART = /^xl\/charts\/chartEx(\d+)\.xml$/i;

/** True when `type` is one of the Office 2016+ ChartEx (`cx:chartSpace`) families. */
export const isChartExViewType = (type: ChartViewModel['type']): boolean =>
	CHART_EX_VIEW_TYPES.has(type);

/**
 * Fallback for embedded workbooks whose ChartEx part is not anchored on a worksheet drawing
 * (for example, a chartsheet-hosted `xl/charts/chartEx*.xml`). Reads cached series values and
 * theme colours through the shared `ooxml-core/chart` parser and view builder.
 */
async function readStandaloneChartEx(bytes: Uint8Array): Promise<ChartViewModel | undefined> {
	let zip: JSZip;
	try {
		zip = await openZip(bytes);
	} catch {
		return undefined;
	}
	const parts = Object.keys(zip.files)
		.map((name) => ({ name, match: CHART_EX_PART.exec(name) }))
		.filter((entry): entry is { name: string; match: RegExpExecArray } => entry.match !== null)
		.sort((a, b) => Number(a.match[1]) - Number(b.match[1]));
	if (parts.length === 0) return undefined;
	let palette = chartThemePalette(undefined);
	const themeEntry = zip.file('xl/theme/theme1.xml');
	if (themeEntry) {
		try {
			const theme = parseThemeXml(await themeEntry.async('string'));
			palette = chartThemePalette(theme.colorScheme, {
				...(theme.fontScheme.major.latin ? { major: theme.fontScheme.major.latin } : {}),
				...(theme.fontScheme.minor.latin ? { minor: theme.fontScheme.minor.latin } : {}),
			});
		} catch {
			// Fall back to the default Office chart theme palette.
		}
	}
	for (const { name } of parts) {
		try {
			const xml = await zip.file(name)?.async('string');
			if (!xml) continue;
			const parsed = parseChartExSpace(xml);
			const extended = chartExSummaryFromModel(parsed.chartExSpace);
			if (extended.summary) return chartSummaryView(extended.summary, palette);
		} catch {
			// Try the next ChartEx part if one is malformed.
		}
	}
	return undefined;
}

/**
 * An embedded Excel workbook that holds a chart (Insert > Chart in Visio, or this editor): the
 * Foreign shape's ForeignData is an Object whose relationship is a package. Anything else embedded
 * (a Word document, a PDF) is not read and the shape is reported as unsupported, as before.
 */
export async function readEmbeddedChartObject(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
): Promise<VisioEmbeddedObject | undefined> {
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(sourcePart)).get(id) : undefined;
	if (!rel || rel.mode !== 'Internal' || rel.type !== `${NS.r}/package`) return undefined;
	if (pkg.getPartByteLength(rel.target) > VISIO_CHART_OBJECT_MAX_BYTES) {
		report('embedded-object-limit', 'An embedded object is larger than 8 MiB and was not read.');
		return undefined;
	}
	const bytes = await pkg.readBytes(rel.target);
	const reader = await import('#embedded-workbook-chart').catch(() => undefined);
	const embedded = await reader?.readEmbeddedChart(bytes);
	const view = embedded?.view ?? (await readStandaloneChartEx(bytes));
	if (!view) return undefined;
	const chartEx =
		!embedded ||
		Boolean(embedded.chart.partName && CHART_EX_PART.test(embedded.chart.partName)) ||
		isChartExViewType(view.type);
	const objectType = Number(foreignData.getAttribute('ObjectType'));
	return {
		...(Number.isFinite(objectType) ? { objectType } : {}),
		bytes,
		chart: view,
		...(chartEx ? { chartEx: true } : {}),
	};
}
