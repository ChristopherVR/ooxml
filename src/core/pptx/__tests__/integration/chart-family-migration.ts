import JSZip from 'jszip';

import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxElement } from '../../index';
import { canonicalXmlLines, lineDiff } from './chart-save-parity-canon';
import type { SaveDeck } from './chart-save-parity';

/**
 * Old-versus-new part diff for one PowerPoint chart source adapter moving onto the neutral writer
 * (docs/agnostic-core-plan.md, chart step C3): the same model edit is saved by `PptxHandler.save`
 * twice, once through the neutral plan (the default) and once with the load-time baseline removed,
 * which is what sends a part to the object-tree writers. Both parts are reloaded and compared.
 */

export interface FamilyEdit {
	name: string;
	/** Whether the edit applies to the chart as loaded. */
	applies: (chart: PptxChartData) => boolean;
	edit: (chart: PptxChartData) => void;
	/** The model field the edit sets (compared after reload). */
	read: (chart: PptxChartData) => unknown;
}

type WithRuntime = { runtime: { chartDataBaselines: Map<string, string> } };
const baselines = (handler: PptxHandler) =>
	(handler as unknown as WithRuntime).runtime.chartDataBaselines;

const toBuffer = (bytes: Uint8Array): ArrayBuffer => new Uint8Array(bytes).buffer;

async function loadCharts(bytes: ArrayBuffer) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes);
	const charts = new Map<string, { element: ChartPptxElement; slide: { isDirty?: boolean } }>();
	for (const slide of data.slides) {
		const visit = (list: PptxElement[] | undefined) => {
			for (const element of list ?? []) {
				const chart = element as ChartPptxElement;
				if (element.type === 'chart' && chart.chartData?.chartPartPath)
					charts.set(chart.chartData.chartPartPath, { element: chart, slide });
				visit((element as { children?: PptxElement[] }).children);
			}
		};
		visit(slide.elements);
	}
	return { handler, data, charts };
}

export interface FamilySave {
	part: string;
	xml: string;
	chart: PptxChartData;
}

/** Saves `edit` on every applicable chart of `deck`, through the neutral plan or the object tree. */
export async function saveFamilyEdit(
	deck: SaveDeck,
	edit: FamilyEdit,
	path: 'neutral' | 'legacy',
): Promise<FamilySave[]> {
	const loaded = await loadCharts(toBuffer(deck.bytes));
	const parts: string[] = [];
	for (const [part, target] of loaded.charts) {
		const chart = target.element.chartData!;
		if (!/<c:chartSpace[\s>]/u.test(deck.parts.find((p) => p.part === part)?.xml ?? '')) continue;
		if (!edit.applies(chart)) continue;
		edit.edit(chart);
		target.slide.isDirty = true;
		parts.push(part);
		// Without its load-time baseline a part has no neutral plan, as for an unclaimed edit.
		if (path === 'legacy') baselines(loaded.handler).delete(part);
	}
	if (parts.length === 0) return [];
	const saved = await loaded.handler.save(loaded.data.slides);
	const zip = await JSZip.loadAsync(saved);
	const reloaded = await loadCharts(toBuffer(saved));
	return Promise.all(
		parts.map(async (part) => ({
			part,
			xml: await zip.file(part)!.async('string'),
			chart: reloaded.charts.get(part)!.element.chartData!,
		})),
	);
}

/** Canonical XML lines only one of the two saves wrote (`-` legacy, `+` neutral). */
export const familyXmlDifferences = (legacy: string, neutral: string): string[] =>
	lineDiff(canonicalXmlLines(legacy), canonicalXmlLines(neutral));
