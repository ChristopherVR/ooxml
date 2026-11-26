import JSZip from 'jszip';

import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxElement, PptxSlide } from '../../index';

interface LoadedChart {
	element: ChartPptxElement & { chartData: PptxChartData };
	slide: PptxSlide;
}

/** Each save starts with a fresh session because it can mutate the model and handler. */
export async function loadChartDeck(bytes: Uint8Array) {
	const handler = new PptxHandler();
	// Copy only the byte view: a Node Buffer can be backed by a larger shared pool.
	const data = await handler.load(new Uint8Array(bytes).buffer);
	const charts = new Map<string, LoadedChart>();
	for (const slide of data.slides) {
		const visit = (elements: readonly PptxElement[]) => {
			for (const element of elements) {
				if (hasChartData(element) && element.chartData.chartPartPath) {
					charts.set(element.chartData.chartPartPath, { element, slide });
				}
				if (element.type === 'group') visit(element.children);
			}
		};
		visit(slide.elements);
	}
	return { charts, handler, data };
}

function hasChartData(
	element: PptxElement,
): element is ChartPptxElement & { chartData: PptxChartData } {
	return element.type === 'chart' && element.chartData !== undefined;
}

/** Share the source entries, with a fresh file table for callers to replace parts. */
export function copyChartArchive(source: JSZip): JSZip {
	const zip = new JSZip();
	Object.assign(zip.files, source.files);
	if ('comment' in source) Object.assign(zip, { comment: source.comment });
	return zip;
}

export interface SavedChart {
	xml: string;
	chart: PptxChartData;
}

/** Read the XML from the output package and the chart model from a fresh product load. */
export async function readSavedCharts(bytes: Uint8Array, parts: readonly string[]) {
	const zip = await JSZip.loadAsync(bytes);
	const { charts, handler } = await loadChartDeck(bytes);
	try {
		const out = new Map<string, SavedChart>();
		for (const part of parts)
			out.set(part, {
				xml: await zip.file(part)!.async('string'),
				chart: charts.get(part)!.element.chartData,
			});
		return out;
	} finally {
		handler.dispose();
	}
}

export interface ChartFixture<Space> {
	readonly name: string;
	readonly bytes: Uint8Array;
	readonly parts: readonly { part: string; space: Space }[];
	readonly archive: JSZip;
}

export interface ChartFixtureEdit<Space> {
	/** Whether the edit applies to the chart as loaded. */
	applies?: (space: Space, chart: PptxChartData) => boolean;
	pptx: (element: ChartPptxElement) => void;
	/** The reference receives a copy of the model before the PowerPoint edit. */
	neutral: (space: Space, chart: PptxChartData) => Space;
}

/** Save every applicable part through PowerPoint and through its neutral reference writer. */
export async function compareChartSaves<Space>(
	deck: ChartFixture<Space>,
	edit: ChartFixtureEdit<Space>,
	write: (space: Space) => string,
) {
	const loaded = await loadChartDeck(deck.bytes);
	try {
		const targets = deck.parts.flatMap((part) => {
			const target = loaded.charts.get(part.part);
			if (!target || (edit.applies?.(part.space, target.element.chartData) ?? true) === false)
				return [];
			return [{ ...part, target, chart: structuredClone(target.element.chartData) }];
		});
		if (targets.length === 0) return undefined;
		const zip = copyChartArchive(deck.archive);
		for (const { part, space, chart, target } of targets) {
			edit.pptx(target.element);
			target.slide.isDirty = true;
			zip.file(part, write(edit.neutral(space, chart)));
		}
		const parts = targets.map(({ part }) => part);
		const [fromPptx, fromNeutral] = await Promise.all([
			loaded.handler.save(loaded.data.slides).then((bytes) => readSavedCharts(bytes, parts)),
			zip
				.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
				.then((bytes) => readSavedCharts(bytes, parts)),
		]);
		return { parts, fromPptx, fromNeutral };
	} finally {
		loaded.handler.dispose();
	}
}
