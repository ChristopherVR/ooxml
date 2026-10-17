import JSZip from 'jszip';
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import {
	chartExSummary,
	parseChartSpace,
	readChartFormatting,
	renderChartSpaceSvg,
	renderChartSummarySvg,
} from '../../../chart/index';
import { parseXml } from '../../../xml/index';

/**
 * Chart renderer gap harness (docs/chart-renderer-gap-report.md). Every chart part of every
 * committed deck runs through core's own pipeline (`parseChartSpace`, summary, `renderChartSvg`).
 * Each feature the PowerPoint renderer draws is then ablated: the feature is stripped from the
 * part text and the SVG is rendered again; when the two SVGs are identical core does not draw it.
 * The harness asserts that nothing throws, writes the table when `CHART_GAP_REPORT` names a file and
 * the SVGs when `CHART_GAP_SVG_DIR` names a folder.
 */

const ROOT = path.resolve(__dirname, '../fixtures');
const CHART_PART = /^ppt\/charts\/chart[^/]*\.xml$/u;
const SIZE = { width: 640, height: 400 };

function walk(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true })
		.flatMap((entry) => {
			const full = path.join(dir, entry.name);
			if (entry.isDirectory()) return walk(full);
			return entry.name.endsWith('.pptx') ? [full] : [];
		})
		.sort();
}

/** Removes `<c:tag>...</c:tag>` (or a self-closing one) from the part text. */
function strip(xml: string, tag: string): string {
	return xml
		.replace(new RegExp(`<c:${tag}(?:\\s[^>]*)?/>`, 'gu'), '')
		.replace(new RegExp(`<c:${tag}(?:\\s[^>]*)?>[\\s\\S]*?</c:${tag}>`, 'gu'), '');
}

interface Feature {
	id: string;
	/** The part uses the feature. */
	uses: (xml: string) => boolean;
	/** The part with the feature removed; absent when the harness cannot ablate it. */
	without?: (xml: string) => string;
}

const has = (re: RegExp) => (xml: string) => re.test(xml);
const count = (re: RegExp, xml: string) => (xml.match(re) ?? []).length;
const FEATURES: Feature[] = [
	{
		id: 'data labels',
		uses: has(/<c:show(?:Val|CatName|Percent|SerName)\s+val="1"/u),
		without: (x) => strip(x, 'dLbls'),
	},
	{ id: 'data label positions (dLblPos)', uses: has(/<c:dLblPos/u) },
	{ id: 'per-point data labels (c:dLbl)', uses: has(/<c:dLbl>/u) },
	{
		id: 'leader lines / callout labels',
		uses: has(/leaderLines|showLeaderLines|wedgeRectCallout|c15:spPr/u),
	},
	{ id: 'data table', uses: has(/<c:dTable/u), without: (x) => strip(x, 'dTable') },
	{ id: 'secondary axis', uses: (x) => count(/<c:valAx>/gu, x) > 1 },
	{
		id: 'combo (several chart groups)',
		uses: (x) =>
			count(/<c:(?:bar|line|area|scatter|pie|doughnut|bubble|radar|stock)Chart>/gu, x) > 1,
	},
	{ id: 'date axis', uses: has(/<c:dateAx/u) },
	{ id: 'multi-level categories', uses: has(/<c:multiLvlStrRef/u) },
	{ id: 'gradient fills', uses: has(/<a:gradFill/u) },
	{ id: 'picture fills', uses: has(/<a:blipFill/u) },
	{ id: 'pattern fills', uses: has(/<a:pattFill/u) },
	{
		id: 'major gridlines',
		uses: has(/<c:majorGridlines/u),
		without: (x) => strip(x, 'majorGridlines'),
	},
	{
		id: 'minor gridlines',
		uses: has(/<c:minorGridlines/u),
		without: (x) => strip(x, 'minorGridlines'),
	},
	{ id: 'axis titles', uses: has(/<c:(?:valAx|catAx|dateAx)>[\s\S]*?<c:title>/u) },
	{ id: 'log scale', uses: has(/<c:logBase/u), without: (x) => strip(x, 'logBase') },
	{ id: 'display units', uses: has(/<c:dispUnits/u), without: (x) => strip(x, 'dispUnits') },
	{ id: 'manual layout', uses: has(/<c:manualLayout/u), without: (x) => strip(x, 'manualLayout') },
	{ id: 'trendlines', uses: has(/<c:trendline>/u), without: (x) => strip(x, 'trendline') },
	{ id: 'error bars', uses: has(/<c:errBars>/u), without: (x) => strip(x, 'errBars') },
	{ id: 'markers', uses: has(/<c:marker>/u), without: (x) => strip(x, 'marker') },
	{ id: 'up/down bars, high-low lines', uses: has(/<c:upDownBars|<c:hiLowLines/u) },
	{
		id: '3-D view',
		uses: has(/<c:view3D|<c:(?:bar|pie|line|area|surface)3DChart/u),
		without: (x) => strip(x, 'view3D'),
	},
	{ id: 'surface', uses: has(/<c:surface3?DChart/u) },
	{ id: 'of-pie', uses: has(/<c:ofPieChart/u) },
	{ id: 'doughnut', uses: has(/<c:doughnutChart/u) },
	{ id: 'explosion', uses: has(/<c:explosion/u), without: (x) => strip(x, 'explosion') },
	{ id: 'smooth lines', uses: has(/<c:smooth val="1"/u) },
	{
		id: 'legend manual layout / entries',
		uses: has(/<c:legend>[\s\S]*?(?:<c:manualLayout|<c:legendEntry)/u),
	},
	{ id: 'chart title overlay', uses: has(/<c:title>[\s\S]*?<c:overlay val="1"/u) },
	{ id: 'rounded corners', uses: has(/<c:roundedCorners val="1"/u) },
	{ id: 'axis number format', uses: has(/<c:numFmt [^>]*sourceLinked="0"/u) },
	{ id: 'text properties (c:txPr)', uses: has(/<c:txPr>/u), without: (x) => strip(x, 'txPr') },
];

interface Row {
	deck: string;
	part: string;
	kind: string;
	chars: number;
	error?: string;
	/** Features the part uses. */
	uses: string[];
	/** Features whose removal changed the SVG (core draws them). */
	drawn: string[];
	userShapes: boolean;
}

function render(xml: string): { svg: string; kind: string } {
	if (/<cx:chartSpace/u.test(xml)) {
		const ex = chartExSummary(xml);
		if (!ex) return { svg: '', kind: 'chartex:unreadable' };
		// Funnel, box and whisker, region map and pareto have no summary: core draws nothing.
		if (!ex.summary) return { svg: '', kind: `chartex:${ex.layoutId} (not drawn)` };
		return { svg: renderChartSummarySvg(ex.summary, SIZE), kind: `chartex:${ex.layoutId}` };
	}
	const { chartSpace } = parseChartSpace(xml);
	const kind = chartSpace.plotArea.groups.map((g) => g.kind).join('+') || 'none';
	// The same feed Word uses: the part's direct formatting beside the model.
	const formatting = chartSpace.plotArea.groups.length
		? readChartFormatting(parseXml(xml, { label: 'Chart part' }).documentElement)
		: undefined;
	return {
		svg: renderChartSpaceSvg(chartSpace, { ...SIZE, ...(formatting ? { formatting } : {}) }),
		kind,
	};
}

async function collect(): Promise<Row[]> {
	const rows: Row[] = [];
	for (const file of walk(ROOT)) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(file));
		} catch {
			continue;
		}
		const deck = path.relative(ROOT, file).replace(/\\/gu, '/');
		const parts = Object.keys(zip.files)
			.filter((name) => CHART_PART.test(name))
			.sort();
		for (const part of parts) {
			const xml = await zip.file(part)!.async('string');
			const rels = zip.file(part.replace('charts/', 'charts/_rels/') + '.rels');
			const relXml = rels ? await rels.async('string') : '';
			const row: Row = {
				deck,
				part,
				kind: '',
				chars: 0,
				uses: [],
				drawn: [],
				userShapes: /chartUserShapes/u.test(relXml),
			};
			try {
				const base = render(xml);
				row.kind = base.kind;
				const svgDir = process.env['CHART_GAP_SVG_DIR'];
				if (svgDir && base.svg) {
					mkdirSync(svgDir, { recursive: true });
					const name = `${deck.replace(/[^\w]+/gu, '_')}-${path.basename(part, '.xml')}.svg`;
					writeFileSync(path.join(svgDir, name), base.svg);
				}
				row.chars = base.svg.length;
				for (const feature of FEATURES) {
					if (!feature.uses(xml)) continue;
					row.uses.push(feature.id);
					if (!feature.without || !base.svg) continue;
					if (render(feature.without(xml)).svg !== base.svg) row.drawn.push(feature.id);
				}
			} catch (error) {
				row.error = error instanceof Error ? error.message : String(error);
			}
			rows.push(row);
		}
	}
	return rows;
}

const ROWS = await collect();

describe('chart renderer gap harness', () => {
	it('renders every chart part of the corpus through core without throwing', () => {
		expect(ROWS.filter((r) => r.error).map((r) => `${r.deck} ${r.part}: ${r.error}`)).toEqual([]);
		expect(ROWS).toHaveLength(82);
	});

	it('reports the features core does not draw', () => {
		const out = process.env['CHART_GAP_REPORT'];
		const table = FEATURES.map((f) => {
			const users = ROWS.filter((r) => r.uses.includes(f.id));
			const drawn = users.filter((r) => r.drawn.includes(f.id)).length;
			return {
				feature: f.id,
				parts: users.length,
				decks: new Set(users.map((r) => r.deck)).size,
				coreDraws: f.without ? `${drawn}/${users.length}` : 'not ablated',
			};
		});
		const kinds: Record<string, number> = {};
		for (const r of ROWS) kinds[r.kind] = (kinds[r.kind] ?? 0) + 1;
		const userShapes = ROWS.filter((r) => r.userShapes).length;
		if (out) {
			mkdirSync(path.dirname(out), { recursive: true });
			writeFileSync(
				out,
				JSON.stringify({ table, kinds, userShapes, parts: ROWS.length }, null, '\t'),
			);
		}
		expect(table.find((t) => t.feature === 'data labels')?.parts).toBeGreaterThan(0);
	});
});
