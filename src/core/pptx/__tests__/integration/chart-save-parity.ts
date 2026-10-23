import JSZip from 'jszip';
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { parseChartSpace, writeChartSpace, type ChartSpace } from '../../../chart/index';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxElement } from '../../index';
import { canonicalXmlLines, lineDiff } from './chart-save-parity-canon';
import type { SaveEdit } from './chart-save-parity-edits';

/**
 * Support for the chart save harness (`chart-save-parity.test.ts`): load every committed deck that
 * holds a classic chart part, make one edit on every chart it applies to, and save it two ways:
 * through PowerPoint's save (`PptxHandler.save`, the object-tree writers plus whatever already goes
 * through `writeChartSpace`) and through the neutral reference (`parseChartSpace` of the part, the
 * equivalent `chart/edit` operation, `writeChartSpace`). Both results are compared as canonical
 * XML and as the `PptxChartData` they load back as.
 */

const ROOT = path.resolve(__dirname, '../fixtures');
const CHART_PART = /^ppt\/charts\/chart[^/]*\.xml$/u;

function walk(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true })
		.flatMap((entry) => {
			const full = path.join(dir, entry.name);
			if (entry.isDirectory()) return walk(full);
			return entry.name.endsWith('.pptx') ? [full] : [];
		})
		.sort();
}

export interface ClassicPart {
	part: string;
	xml: string;
	space: ChartSpace;
	/** Whether `parseChartSpace` reported anything it does not model (the writer would drop it). */
	clean: boolean;
}

export interface SaveDeck {
	name: string;
	bytes: Buffer;
	parts: ClassicPart[];
}

// A copy through the constructor: `slice()` on a Node Buffer (`readFileSync`) is a view, not a copy,
// and since Node 24.21 (`Buffer.poolSize` 64 KiB) a file under 32 KiB is a view into a shared pool.
const toBuffer = (bytes: Uint8Array): ArrayBuffer => new Uint8Array(bytes).buffer;

/** Every committed deck with at least one classic (`c:chartSpace`) chart part. */
export async function classicChartDecks(): Promise<SaveDeck[]> {
	const decks: SaveDeck[] = [];
	for (const file of walk(ROOT)) {
		const bytes = readFileSync(file);
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(bytes);
		} catch {
			continue;
		}
		const parts: ClassicPart[] = [];
		for (const part of Object.keys(zip.files).filter((name) => CHART_PART.test(name))) {
			const xml = await zip.file(part)!.async('string');
			if (!/<c:chartSpace[\s>]/u.test(xml)) continue;
			const { chartSpace, issues } = parseChartSpace(xml);
			parts.push({ part, xml, space: chartSpace, clean: issues.length === 0 });
		}
		parts.sort((a, b) => a.part.localeCompare(b.part));
		if (parts.length > 0)
			decks.push({ name: path.relative(ROOT, file).replace(/\\/gu, '/'), bytes, parts });
	}
	return decks;
}

interface LoadedChart {
	element: ChartPptxElement;
	slide: { isDirty?: boolean };
}

async function load(bytes: ArrayBuffer) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes);
	const charts = new Map<string, LoadedChart>();
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
	return { charts, handler, data };
}

export interface SavedChart {
	xml: string;
	chart: PptxChartData;
}

async function readBack(bytes: Uint8Array, parts: readonly string[]) {
	const zip = await JSZip.loadAsync(bytes);
	const loaded = await load(toBuffer(bytes));
	const out = new Map<string, SavedChart>();
	for (const part of parts)
		out.set(part, {
			xml: await zip.file(part)!.async('string'),
			chart: loaded.charts.get(part)!.element.chartData!,
		});
	return out;
}

export interface EditTarget extends ClassicPart {
	/** The chart as PowerPoint loaded it, before the edit. */
	chart: PptxChartData;
}

/** The parts of `deck` the edit applies to. */
export async function applicableParts(deck: SaveDeck, edit: SaveEdit): Promise<EditTarget[]> {
	const loaded = await load(toBuffer(deck.bytes));
	return deck.parts.flatMap((part) => {
		const chart = loaded.charts.get(part.part)?.element.chartData;
		return chart !== undefined && (edit.applies?.(part.space, chart) ?? true)
			? [{ ...part, chart: structuredClone(chart) }]
			: [];
	});
}

/** PowerPoint's save: edit every applicable chart's model, save the deck, read the parts back. */
export async function savedByPptx(deck: SaveDeck, edit: SaveEdit, parts: readonly EditTarget[]) {
	const loaded = await load(toBuffer(deck.bytes));
	for (const { part } of parts) {
		const target = loaded.charts.get(part)!;
		edit.pptx(target.element);
		target.slide.isDirty = true;
	}
	const saved = await loaded.handler.save(loaded.data.slides);
	return readBack(
		saved,
		parts.map((item) => item.part),
	);
}

/** The neutral reference: edit each parsed part, write it, load the deck. */
export async function savedByNeutral(deck: SaveDeck, edit: SaveEdit, parts: readonly EditTarget[]) {
	const zip = await JSZip.loadAsync(deck.bytes);
	for (const { part, space, chart } of parts)
		zip.file(part, writeChartSpace(edit.neutral(space, chart)));
	return readBack(
		await zip.generateAsync({ type: 'uint8array' }),
		parts.map((item) => item.part),
	);
}

/**
 * What depends on the embedded workbook: PowerPoint's save writes edited series data into it, the
 * neutral reference writes the chart part only, so the two are not compared. That covers the
 * workbook contents and whether data labels may be edited through it (`dataLabelTextEditable`
 * and its `dataLabelSourceShape`, set only while every referenced cell agrees with the caches).
 */
const WORKBOOK_PATHS = new Set([
	'.embeddedWorkbookData',
	'.externalData.embeddedWorkbookData',
	'.dataLabelTextEditable',
	'.dataLabelSourceShape',
]);

/** Paths where two JSON values differ, with both values. */
export function modelDifferences(a: unknown, b: unknown, at = ''): string[] {
	if (WORKBOOK_PATHS.has(at) || JSON.stringify(a) === JSON.stringify(b)) return [];
	if (a && b && typeof a === 'object' && typeof b === 'object') {
		const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
		return [...keys].flatMap((key) =>
			modelDifferences(
				(a as Record<string, unknown>)[key],
				(b as Record<string, unknown>)[key],
				`${at}.${key}`,
			),
		);
	}
	return [`${at}: pptx ${JSON.stringify(a)}, neutral ${JSON.stringify(b)}`];
}

/** Canonical XML difference of two saved parts: `-` only PowerPoint wrote, `+` only neutral. */
export const xmlDifferences = (pptx: string, neutral: string): string[] =>
	lineDiff(canonicalXmlLines(pptx), canonicalXmlLines(neutral));
