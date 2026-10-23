// Inserting and editing charts and pictures. A chart is edited over its neutral model
// (`ChartObject.chartSpace`, docs/agnostic-core-plan.md chart step C2): `editChart` applies a
// `chart/edit` intent, `updateChart` maps a summary patch onto the same operations
// (`chart-space-patch.ts`), and the summary the painter reads is derived from the edited model.
import type { ChartObject, DrawingAnchor, ImageObject, Workbook } from '../model';
import { type EditContext, sheetAt } from './context';
import { assertBarClusterOptions } from '../../chart/bar-cluster-geometry';
import { applyChartEditIntent, type ChartEditIntent } from '../../chart/edit/intent';
import { parseChartSpace } from '../../chart/parse-space';
import type { ChartSpace } from '../../chart/model';
import { chartWithElementFills, type ChartElementFills } from './chart-element-fill';
import { chartSpaceWithPatch } from './chart-space-patch';
import { chartObjectFromSpace } from '../read/chart';
import { newChartSpaceOf } from '../write/chart';

/** File extensions for the picture types Excel accepts. */
export const IMAGE_EXTENSIONS: Readonly<Record<string, string>> = {
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/tiff': 'tiff',
	'image/webp': 'webp',
	'image/svg+xml': 'svg',
	'image/x-emf': 'emf',
	'image/x-wmf': 'wmf',
};

/** What `updateChart` may change (the chart keeps its kind, source part and model). */
export type ChartPatch = Partial<
	Omit<ChartObject, 'kind' | 'partName' | 'styleDefinition' | 'formatting' | 'chartSpace'>
> & { elementFills?: ChartElementFills };

/**
 * The model a chart is edited over: its own, else its kept part parsed, else the part a new chart
 * is written as.
 */
export function chartSpaceOf(workbook: Workbook, chart: ChartObject): ChartSpace {
	if (chart.chartSpace) return chart.chartSpace;
	const bytes = chart.partName ? workbook.source?.parts.get(chart.partName) : undefined;
	if (bytes) return parseChartSpace(new TextDecoder().decode(bytes)).chartSpace;
	return newChartSpaceOf(chart);
}

/**
 * Adds a chart (a new chart part is written on save); returns its index. A chart that carries a
 * `chartSpace` (a copy of another chart) keeps it; otherwise the part is built from the summary.
 */
export function addChart(ctx: EditContext, s: number, chart: Omit<ChartObject, 'kind'>): number {
	assertBarClusterOptions(chart);
	const sheet = sheetAt(ctx.workbook, s);
	const base: ChartObject = { ...structuredClone(chart), kind: 'chart' };
	delete base.partName;
	const added = chartObjectFromSpace(base.chartSpace ?? newChartSpaceOf(base), base);
	return ctx.run(
		'Insert chart',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.drawings.push(structuredClone(added));
			return sheet.drawings.length - 1;
		},
		{ sheet: s },
	);
}

/** Replaces the chart at `index` with `next` as one undoable step. */
function replaceChart(
	ctx: EditContext,
	s: number,
	index: number,
	label: string,
	next: ChartObject,
) {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		label,
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const live = sheet.drawings[index] as ChartObject;
			for (const key of Object.keys(live)) Reflect.deleteProperty(live, key);
			Object.assign(live, structuredClone(next));
		},
		{ sheet: s },
	);
}

function chartAt(ctx: EditContext, s: number, index: number): ChartObject {
	const chart = sheetAt(ctx.workbook, s).drawings[index];
	if (!chart || chart.kind !== 'chart') throw new RangeError(`No chart at index ${index}`);
	return chart;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** The summary `patch` describes, merged over `chart` as `updateChart` always has. */
function patchedSummary(chart: ChartObject, patch: ChartPatch): ChartObject {
	const after = structuredClone(chart);
	const changes = structuredClone(patch);
	delete changes.elementFills;
	// A legacy color edit replaces the parsed DrawingML choice unless the caller edits both.
	changes.series?.forEach((series, i) => {
		const before = after.series[i];
		if (!before) return;
		const legacyChanged = !same(before.color, series.color);
		const drawingChanged = !same(before.drawingColor, series.drawingColor);
		if (legacyChanged && !drawingChanged) delete series.drawingColor;
		if ((legacyChanged || drawingChanged) && same(before.fill, series.fill)) delete series.fill;
		for (const key of new Set([
			...Object.keys(before.pointColors ?? {}),
			...Object.keys(series.pointColors ?? {}),
		])) {
			const point = Number(key);
			if (
				!same(before.pointColors?.[point], series.pointColors?.[point]) &&
				same(before.pointFills?.[point], series.pointFills?.[point])
			)
				delete series.pointFills?.[point];
		}
	});
	for (const [key, value] of Object.entries(changes)) {
		if (['kind', 'partName', 'chartSpace', '__proto__', 'constructor', 'prototype'].includes(key))
			continue;
		if (value === undefined) Reflect.deleteProperty(after, key);
		else Reflect.set(after, key, value);
	}
	return after;
}

/**
 * Edits a chart through its summary fields. The change is applied to the chart's model with the
 * `chart/edit` operations (`chart-space-patch.ts`), so a part read from a file keeps everything the
 * patch does not name; on save an edited model is written with `chart/writeChartSpace` and an
 * unedited one keeps its part byte for byte.
 */
export function updateChart(ctx: EditContext, s: number, index: number, patch: ChartPatch): void {
	assertBarClusterOptions(patch);
	const chart = chartAt(ctx, s, index);
	if (patch.elementFills) chartWithElementFills(chart, patch.elementFills);
	const space = chartSpaceOf(ctx.workbook, chart);
	// The summary the model shows; the live summary with the patch is what it should show, so a
	// change made to the summary fields directly is applied with the patch.
	const before = chartObjectFromSpace(space, chart);
	const after = patchedSummary(chart, patch);
	const next = chartSpaceWithPatch(space, before, after, {
		...(patch.quickLayout === undefined ? {} : { layout: patch.quickLayout }),
		...(patch.elementFills ? { fills: patch.elementFills } : {}),
	});
	replaceChart(ctx, s, index, 'Edit chart', chartObjectFromSpace(next, after));
}

/**
 * Applies a `chart/edit` intent (what a Format pane sends) to a chart's model. A Change Colors
 * intent also records the new palette, which the colour style part is written from.
 */
export function editChart(
	ctx: EditContext,
	s: number,
	index: number,
	intent: ChartEditIntent,
): void {
	const chart = chartAt(ctx, s, index);
	const space = chartSpaceOf(ctx.workbook, chart);
	const next = applyChartEditIntent(space, intent);
	if (same(next, space)) return;
	const frame = intent.op === 'colorPalette' ? { ...chart, colorPalette: intent.to } : chart;
	replaceChart(ctx, s, index, 'Format chart', chartObjectFromSpace(next, frame));
}

/** A media part name not used by the workbook's package yet (`xl/media/imageN.ext`). */
export function newMediaPart(workbook: Workbook, contentType: string): string {
	const ext = IMAGE_EXTENSIONS[contentType] ?? contentType.split('/').pop() ?? 'bin';
	const taken = new Set([...(workbook.source?.parts.keys() ?? [])].map((p) => p.toLowerCase()));
	for (const sheet of workbook.sheets)
		for (const d of sheet.drawings) if (d.kind === 'image') taken.add(d.partName.toLowerCase());
	for (let n = 1; ; n++) {
		const name = `xl/media/image${n}.${ext}`;
		if (!taken.has(name)) return name;
	}
}

/**
 * Inserts a picture. Its bytes are stored in the workbook's package (created when the workbook
 * has none) under a new media part, which the writer copies on save. Returns the image.
 */
export function addImage(
	ctx: EditContext,
	s: number,
	bytes: Uint8Array,
	contentType: string,
	anchor: DrawingAnchor,
	name?: string,
): ImageObject {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	if (!contentType.startsWith('image/')) throw new Error(`Not an image type: ${contentType}`);
	const partName = newMediaPart(workbook, contentType);
	return ctx.run(
		'Insert picture',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			workbook.source ??= { parts: new Map() };
			workbook.source.parts.set(partName, bytes.slice());
			const image: ImageObject = {
				kind: 'image',
				anchor: structuredClone(anchor),
				partName,
				contentType,
			};
			if (name) image.name = name;
			sheet.drawings.push(image);
			return structuredClone(image);
		},
		{ sheet: s },
	);
}
