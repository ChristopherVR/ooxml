import type {
	PptxElement,
	PptxSlide,
	PptxSmartArtNode,
	PptxTableCell,
	TextSegment,
} from '../../index';
import { hasTextProperties } from '../../index';
import { chartLabelsAreEditable } from './chart-label-editability';

export interface FindTextTarget {
	elementId: string;
	notes?: boolean;
	elementPath?: number[];
	tableCell?: { rowIndex: number; colIndex: number };
	smartArtNodeId?: string;
	chartTitle?: boolean;
	chartSeriesIndex?: number;
	chartCategoryIndex?: number;
	readOnly?: boolean;
	runs: readonly FindTextRun[];
}
export interface FindTextRun {
	text: string;
	isParagraphBreak?: boolean;
	isLineBreak?: boolean;
}

export function* smartArtNodes(nodes: readonly PptxSmartArtNode[]): Generator<PptxSmartArtNode> {
	for (const node of nodes) {
		yield node;
		if (node.children) yield* smartArtNodes(node.children);
	}
}

export function textRuns(cell: PptxTableCell): readonly FindTextRun[] {
	return cell.textRuns?.length ? cell.textRuns : [{ text: cell.text }];
}

/** Walk text-bearing objects while retaining the top-level selection target for the UI. */
export function* findTextTargets(
	element: PptxElement,
	elementId = element.id,
	path: number[] = [],
	includeChartTitles = true,
	includeChartData = false,
): Generator<FindTextTarget> {
	const location = { elementId, ...(path.length ? { elementPath: path } : {}) };
	if (element.type === 'group') {
		for (const [index, child] of element.children.entries())
			yield* findTextTargets(
				child,
				elementId,
				[...path, index],
				includeChartTitles,
				includeChartData,
			);
	} else if (hasTextProperties(element)) {
		const runs: readonly TextSegment[] = element.textSegments?.length
			? element.textSegments
			: element.text
				? [{ text: element.text, style: element.textStyle ?? {} }]
				: [];
		yield { ...location, runs };
	} else if (element.type === 'table' && element.tableData) {
		for (const [rowIndex, row] of element.tableData.rows.entries())
			for (const [colIndex, cell] of row.cells.entries())
				yield { ...location, tableCell: { rowIndex, colIndex }, runs: textRuns(cell) };
	} else if (element.type === 'smartArt' && element.smartArtData) {
		for (const node of smartArtNodes(element.smartArtData.nodes))
			yield { ...location, smartArtNodeId: node.id, runs: [{ text: node.text }] };
	} else if (element.type === 'chart' && element.chartData) {
		const chart = element.chartData;
		if (includeChartTitles && chart.style?.hasTitle !== false) {
			const runs = chart.titleRuns?.length
				? chart.titleRuns
				: chart.title
					? [{ text: chart.title }]
					: [];
			yield { ...location, chartTitle: true, readOnly: chart.titleTextEditable === false, runs };
		}
		if (includeChartData) {
			const readOnly = !chartLabelsAreEditable(element);
			for (const [chartSeriesIndex, series] of chart.series.entries())
				yield { ...location, chartSeriesIndex, readOnly, runs: [{ text: series.name }] };
			for (const [chartCategoryIndex, text] of chart.categories.entries())
				yield { ...location, chartCategoryIndex, readOnly, runs: [{ text }] };
		}
	}
}

export function isFindBreak(run: FindTextRun): boolean {
	return Boolean(run.isParagraphBreak || run.isLineBreak || run.text === '\n');
}

/** Replace across formatting boundaries; the replacement inherits its starting run's format. */
export function replaceFindRuns<R extends FindTextRun>(
	runs: readonly R[],
	startIndex: number,
	startOffset: number,
	length: number,
	replacement: string,
): R[] | undefined {
	const first = runs[startIndex];
	if (
		!first ||
		isFindBreak(first) ||
		startOffset < 0 ||
		startOffset >= first.text.length ||
		length <= 0
	)
		return undefined;
	const next = [...runs];
	let remaining = length;
	for (let index = startIndex; index < runs.length && remaining > 0; index++) {
		const run = runs[index]!;
		if (isFindBreak(run)) return undefined;
		const offset = index === startIndex ? startOffset : 0;
		const removed = Math.min(remaining, run.text.length - offset);
		next[index] = {
			...run,
			text:
				run.text.slice(0, offset) +
				(index === startIndex ? replacement : '') +
				run.text.slice(offset + removed),
		};
		remaining -= removed;
	}
	return remaining === 0 ? next : undefined;
}

/** Speaker-note bodies are an explicit search scope; notes-page chrome is excluded. */
export function* findSlideTextTargets(
	slide: PptxSlide,
	includeNotes = false,
	includeChartTitles = true,
	includeChartData = false,
): Generator<FindTextTarget> {
	for (const element of slide.elements)
		yield* findTextTargets(element, element.id, [], includeChartTitles, includeChartData);
	if (includeNotes && slide.notes) {
		const runs =
			slide.notesSegments?.length &&
			slide.notesSegments.map((run) => run.text).join('') === slide.notes
				? slide.notesSegments
				: [{ text: slide.notes, style: {} }];
		yield { elementId: '', notes: true, runs };
	}
}
