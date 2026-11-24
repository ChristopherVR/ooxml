/**
 * Multi-column regions: a run of lines that a vertical gutter (no ink, text on both sides) cuts
 * all the way down is read column by column. Columns must be wide and mostly filled, so a
 * borderless table with wide gaps stays one column (`tables.ts` rebuilds it).
 */
import type { Line } from './lines';
import type { TextArea } from './paragraphs';
import { inkIntervals, splitLine } from './sublines';

export interface SingleBand {
	kind: 'single';
	lines: Line[];
}

export interface ColumnBand {
	kind: 'columns';
	/** Each column's lines, top to bottom. */
	columns: Line[][];
	/** Each column's left and right ink edges. */
	bounds: [number, number][];
	top: number;
	bottom: number;
	/** The narrowest gap between two columns, points. */
	space: number;
}

export type Band = SingleBand | ColumnBand;

const MIN_GUTTER = 8;
/** Lines must cross a gutter this many times for it to count. */
const MIN_STRADDLES = 3;

type Gap = [number, number];

function freeGaps(line: Line, zone: Gap): Gap[] {
	const ink = inkIntervals(line, 4);
	const gaps: Gap[] = [];
	for (let at = 1; at < ink.length; at++) {
		const gap: Gap = [ink[at - 1]![1], ink[at]![0]];
		const middle = (gap[0] + gap[1]) / 2;
		if (gap[1] - gap[0] >= MIN_GUTTER && middle > zone[0] && middle < zone[1]) gaps.push(gap);
	}
	return gaps;
}

/** What is left of a gutter once `line`'s ink is taken out of it (the widest free piece). */
function narrow(gutter: Gap, line: Line): Gap | undefined {
	let pieces: Gap[] = [gutter];
	for (const [start, end] of inkIntervals(line, 0)) {
		pieces = pieces.flatMap(([a, b]): Gap[] => {
			if (end <= a || start >= b) return [[a, b]];
			const parts: Gap[] = [
				[a, Math.min(b, start)],
				[Math.max(a, end), b],
			];
			return parts.filter(([x, y]) => y - x > 0);
		});
	}
	const best = pieces.sort((x, y) => y[1] - y[0] - (x[1] - x[0]))[0];
	return best && best[1] - best[0] >= MIN_GUTTER * 0.75 ? best : undefined;
}

const straddles = (line: Line, gutter: Gap) => {
	const ink = inkIntervals(line, 0);
	return (
		ink.some(([, end]) => end <= gutter[0] + 0.5) && ink.some(([start]) => start >= gutter[1] - 0.5)
	);
};

function region(
	lines: Line[],
	from: number,
	area: TextArea,
): { end: number; band: ColumnBand } | undefined {
	const width = area.right - area.left;
	const zone: Gap = [area.left + 0.15 * width, area.right - 0.15 * width];
	let gutters = freeGaps(lines[from]!, zone);
	if (!gutters.length) return undefined;
	let end = from + 1;
	while (end < lines.length) {
		const line = lines[end]!;
		if (line.baseline - lines[end - 1]!.baseline > 4 * line.size) break;
		const next = gutters.map((gutter) => narrow(gutter, line)).filter((gap): gap is Gap => !!gap);
		if (!next.length) break;
		gutters = next;
		end++;
	}
	const run = lines.slice(from, end);
	gutters = gutters.filter(
		(gutter) => run.filter((line) => straddles(line, gutter)).length >= MIN_STRADDLES,
	);
	if (!gutters.length || run.length < 4) return undefined;
	const cuts = gutters.map(([a, b]) => (a + b) / 2).sort((a, b) => a - b);
	const columns: Line[][] = cuts.map(() => []).concat([[]]);
	for (const line of run)
		splitLine(line, cuts).forEach((piece, index) => {
			if (piece) columns[index]!.push(piece);
		});
	if (columns.some((column) => column.length < 2)) return undefined;
	const bounds = columns.map((column): [number, number] => [
		Math.min(...column.map((line) => line.left)),
		Math.max(...column.map((line) => line.right)),
	]);
	const size = run[0]!.size;
	const widths = bounds.map(([a, b]) => b - a);
	const space = Math.min(...bounds.slice(1).map(([a], index) => a - bounds[index]![1]));
	const filled = columns.every((column, index) => {
		const ratios = column
			.map((line) => (line.right - line.left) / widths[index]!)
			.sort((a, b) => a - b);
		return ratios[Math.floor(ratios.length / 2)]! >= 0.5;
	});
	if (!filled || space <= 0 || Math.min(...widths) < 6 * size || space > 0.5 * Math.min(...widths))
		return undefined;
	return {
		end,
		band: {
			kind: 'columns',
			columns,
			bounds,
			top: Math.min(...run.map((line) => line.top)),
			bottom: Math.max(...run.map((line) => line.bottom)),
			space,
		},
	};
}

/** The page's lines as single-column runs and column regions, top to bottom. */
export function splitColumns(lines: Line[], area: TextArea): Band[] {
	const bands: Band[] = [];
	let single: Line[] = [];
	let at = 0;
	while (at < lines.length) {
		const found = region(lines, at, area);
		if (!found) {
			single.push(lines[at]!);
			at++;
			continue;
		}
		if (single.length) bands.push({ kind: 'single', lines: single });
		single = [];
		bands.push(found.band);
		at = found.end;
	}
	if (single.length) bands.push({ kind: 'single', lines: single });
	return bands;
}
