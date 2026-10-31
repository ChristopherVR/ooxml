import { visioTabStarts, type VisioTabSegment, type VisioTextRun } from 'ooxml-core/visio';
import type { PositionedLine } from './paragraph-layout';
import { measureBrowserText, type MeasureText } from './text-layout';

export interface TabbedSegment {
	runs: VisioTextRun[];
	x: number;
}

/** Whether a line has a tab character to lay out. */
export const lineHasTab = (line: PositionedLine): boolean =>
	line.runs.some((run) => run.text.includes('\t'));

const widthOf = (runs: readonly VisioTextRun[], measure: MeasureText, stop?: string): number => {
	let width = 0;
	for (const run of runs) {
		const at = stop === undefined ? -1 : run.text.indexOf(stop);
		const text = at < 0 ? run.text : run.text.slice(0, at);
		if (text) width += measure(text, run);
		if (at >= 0) break;
	}
	return width;
};

/**
 * Splits a left-aligned line at its tab characters and places each part at the tab stop it
 * reaches (see `visioTabStarts`). The tab characters themselves are not drawn. A centred or
 * right-aligned line is laid out from its own start (Visio 16 measures the stops from there) and
 * the whole line then moves to its anchor; justified and distributed lines keep their single run
 * of text, as before.
 */
export function layoutTabbedLine(
	line: PositionedLine,
	measure: MeasureText = measureBrowserText,
): TabbedSegment[] | undefined {
	const anchored = line.align === 'center' || line.align === 'right';
	if ((line.align !== 'left' && !anchored) || line.direction !== 'ltr' || !lineHasTab(line))
		return undefined;
	const parts: VisioTextRun[][] = [[]];
	for (const run of line.runs) {
		const pieces = run.text.split('\t');
		pieces.forEach((piece, index) => {
			if (index > 0) parts.push([]);
			if (piece) parts[parts.length - 1]!.push({ ...run, text: piece });
		});
	}
	const segments: VisioTabSegment[] = parts.map((runs) => {
		const width = widthOf(runs, measure);
		const decimal = runs.some((run) => run.text.includes('.'))
			? widthOf(runs, measure, '.')
			: width;
		return { width, decimal };
	});
	const stops = line.tabs?.stops ?? [];
	let starts: number[];
	if (anchored) {
		starts = visioTabStarts(segments, stops);
		const last = segments.length - 1;
		const width = starts[last]! + segments[last]!.width;
		const shift = line.x - (line.align === 'center' ? width / 2 : width);
		starts = starts.map((at) => at + shift);
	} else starts = visioTabStarts(segments, stops, line.tabs?.origin ?? 0, line.x);
	return parts.map((runs, index) => ({ runs, x: starts[index]! }));
}
