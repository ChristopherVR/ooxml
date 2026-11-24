/**
 * Vector drawings: the painted paths and the rules that no table, underline or cell shading took
 * become Word drawings. Pieces close to each other form one drawing (so a chart stays one object);
 * each piece keeps its geometry (rectangles, ellipses and lines as preset shapes, other paths as
 * custom geometry), fill, outline, opacity and dashes. A page holds a bounded number of pieces;
 * the rest is reported.
 */
import type { PdfPagePath, PdfPageRule, PdfPathSegment } from '../../pdf/index';
import { isSeparatorRule } from './note-marks';

export type PieceGeometry =
	| { kind: 'rect' | 'ellipse' }
	| { kind: 'line'; flipH: boolean }
	| { kind: 'path'; segments: PdfPathSegment[] };

export interface DrawingPiece {
	/** The geometry's box on the page (points, the outline's width not included). */
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	geometry: PieceGeometry;
	fill?: string;
	stroke?: string;
	lineWidth: number;
	fillAlpha?: number;
	strokeAlpha?: number;
	dash?: number[];
}

/** One drawing at a place on its page: one shape, or several kept together. */
export interface VectorDrawing {
	page: number;
	x: number;
	y: number;
	width: number;
	height: number;
	pieces: DrawingPiece[];
}

/** What a page may turn into drawings, and the shapes it cannot (pattern fills, clips). */
export interface DrawingSource {
	rules: PdfPageRule[];
	paths: PdfPagePath[];
	leftOut: number;
}

/** Pieces a page turns into drawings; more are reported as left out. */
export const MAX_PIECES = 500;
/** Path segments a page's drawings may hold in all. */
export const MAX_SEGMENTS = 20_000;
/** Pieces closer than this (points) belong to one drawing. */
const GAP = 3;

/**
 * The rules of a page that are drawings: not used by a table that was rebuilt (`tables`), not an
 * underline or strike-through (`marks`), and visible. `rest` holds the rules no ruled grid used.
 */
export function drawingSource(
	rules: readonly PdfPageRule[],
	paths: readonly PdfPagePath[],
	rest: readonly PdfPageRule[],
	marks: ReadonlySet<PdfPageRule>,
	tables: readonly { top: number; bottom: number; left: number; columns: number[] }[],
	page: { width: number; height: number },
): DrawingSource {
	const free = new Set(rest);
	const boxes = tables.map((table) => ({
		x0: table.left - 2,
		x1: table.left + table.columns.reduce((sum, width) => sum + width, 0) + 2,
		y0: table.top - 2,
		y1: table.bottom + 2,
	}));
	const inTable = (rule: PdfPageRule) =>
		boxes.some(
			(box) => rule.x0 >= box.x0 && rule.x1 <= box.x1 && rule.y0 >= box.y0 && rule.y1 <= box.y1,
		);
	const out: DrawingSource = { rules: [], paths: [], leftOut: 0 };
	for (const rule of rules) {
		if (marks.has(rule) || (!free.has(rule) && inTable(rule))) continue;
		const w = rule.x1 - rule.x0;
		const h = rule.y1 - rule.y0;
		// A filled line paints nothing; a page-sized white fill is the paper.
		if (rule.kind === 'line' ? !rule.stroked || w + h < 0.5 : w < 0.1 && h < 0.1) continue;
		const paper =
			!rule.stroked && /^#f{6}$/i.test(rule.color) && w * h > page.width * page.height * 0.5;
		if (paper) continue;
		if (rule.pattern || rule.clipped) out.leftOut++;
		else out.rules.push(rule);
	}
	out.paths.push(...paths);
	return out;
}

/** Leaves out the separator rules above the notes that became Word notes (Word draws its own). */
export function withoutNoteSeparators(
	source: DrawingSource,
	noteTop: number | undefined,
	bodySize: number,
	width: number,
): DrawingSource {
	if (noteTop === undefined) return source;
	return {
		...source,
		rules: source.rules.filter(
			(rule) =>
				!(isSeparatorRule(rule, width) && rule.y0 < noteTop && rule.y0 > noteTop - 3 * bodySize),
		),
	};
}

function rulePiece(rule: PdfPageRule): DrawingPiece {
	const base = { x0: rule.x0, y0: rule.y0, x1: rule.x1, y1: rule.y1, lineWidth: rule.lineWidth };
	if (rule.kind === 'line')
		return {
			...base,
			geometry: { kind: 'line', flipH: false },
			stroke: rule.color,
			...(rule.alpha !== undefined ? { strokeAlpha: rule.alpha } : {}),
			...(rule.dash ? { dash: rule.dash } : {}),
		};
	const piece: DrawingPiece = { ...base, geometry: { kind: 'rect' } };
	if (rule.filled) piece.fill = rule.color;
	if (rule.stroked) piece.stroke = rule.strokeColor ?? rule.color;
	if (rule.alpha !== undefined) piece[rule.filled ? 'fillAlpha' : 'strokeAlpha'] = rule.alpha;
	if (rule.stroked && rule.dash) piece.dash = rule.dash;
	return piece;
}

/** Whether the path is an axis-aligned ellipse drawn as four curves through its box's midpoints. */
function isEllipse(path: PdfPagePath): boolean {
	const segments = path.segments.filter((segment) => segment.op !== 'Z');
	const [move, ...curves] = segments;
	if (move?.op !== 'M' || curves.length !== 4 || curves.some((segment) => segment.op !== 'C'))
		return false;
	const w = path.x1 - path.x0;
	const h = path.y1 - path.y0;
	if (w < 1 || h < 1) return false;
	const tolerance = Math.max(0.5, 0.03 * Math.max(w, h));
	const cx = (path.x0 + path.x1) / 2;
	const cy = (path.y0 + path.y1) / 2;
	const middles = [
		[cx, path.y0],
		[path.x1, cy],
		[cx, path.y1],
		[path.x0, cy],
	] as const;
	const ends = curves.map((segment) => segment as { x: number; y: number });
	const last = ends[3]!;
	if (Math.abs(last.x - move.x) > tolerance || Math.abs(last.y - move.y) > tolerance) return false;
	const hit = new Set<number>();
	for (const end of ends) {
		const index = middles.findIndex(
			([x, y]) => Math.abs(end.x - x) <= tolerance && Math.abs(end.y - y) <= tolerance,
		);
		if (index < 0) return false;
		hit.add(index);
	}
	return hit.size === 4;
}

function pathPiece(path: PdfPagePath): DrawingPiece {
	const piece: DrawingPiece = {
		x0: path.x0,
		y0: path.y0,
		x1: path.x1,
		y1: path.y1,
		lineWidth: path.lineWidth,
		geometry: { kind: 'path', segments: path.segments },
	};
	const [first, second, third] = path.segments;
	if (isEllipse(path)) piece.geometry = { kind: 'ellipse' };
	else if (first?.op === 'M' && second?.op === 'L' && !third && !path.fill)
		piece.geometry = { kind: 'line', flipH: (second.x - first.x) * (second.y - first.y) < 0 };
	if (path.fill) piece.fill = path.fill;
	if (path.stroke) piece.stroke = path.stroke;
	if (path.fillAlpha !== undefined) piece.fillAlpha = path.fillAlpha;
	if (path.strokeAlpha !== undefined) piece.strokeAlpha = path.strokeAlpha;
	if (path.dash) piece.dash = path.dash;
	return piece;
}

const segmentCount = (piece: DrawingPiece) =>
	piece.geometry.kind === 'path' ? piece.geometry.segments.length : 5;

/** Groups pieces whose boxes (with their outlines and `GAP`) touch. */
function cluster(pieces: DrawingPiece[]): DrawingPiece[][] {
	const parent = pieces.map((_, index) => index);
	const find = (index: number): number => {
		while (parent[index] !== index) index = parent[index] = parent[parent[index]!]!;
		return index;
	};
	const reach = pieces.map((piece) => GAP / 2 + (piece.stroke ? piece.lineWidth / 2 : 0));
	for (let a = 0; a < pieces.length; a++)
		for (let b = a + 1; b < pieces.length; b++) {
			const p = pieces[a]!;
			const q = pieces[b]!;
			const r = reach[a]! + reach[b]!;
			if (p.x0 - r <= q.x1 && q.x0 - r <= p.x1 && p.y0 - r <= q.y1 && q.y0 - r <= p.y1)
				parent[find(a)] = find(b);
		}
	const groups = new Map<number, DrawingPiece[]>();
	pieces.forEach((piece, index) => {
		const root = find(index);
		const group = groups.get(root);
		if (group) group.push(piece);
		else groups.set(root, [piece]);
	});
	return [...groups.values()];
}

/** The page's drawings, the pieces they hold and the shapes left out over the page's limits. */
export function pageDrawings(
	page: number,
	source: DrawingSource,
): { drawings: VectorDrawing[]; pieces: number; overLimit: number } {
	const all = [...source.rules.map(rulePiece), ...source.paths.map(pathPiece)];
	const pieces: DrawingPiece[] = [];
	let segments = 0;
	let overLimit = 0;
	for (const piece of all) {
		const count = segmentCount(piece);
		if (pieces.length >= MAX_PIECES || segments + count > MAX_SEGMENTS) overLimit++;
		else {
			pieces.push(piece);
			segments += count;
		}
	}
	const drawings = cluster(pieces).map((group): VectorDrawing => {
		const x = Math.min(...group.map((piece) => piece.x0));
		const y = Math.min(...group.map((piece) => piece.y0));
		return {
			page,
			x,
			y,
			width: Math.max(...group.map((piece) => piece.x1)) - x,
			height: Math.max(...group.map((piece) => piece.y1)) - y,
			pieces: group,
		};
	});
	drawings.sort((a, b) => a.y - b.y || a.x - b.x);
	return { drawings, pieces: pieces.length, overLimit };
}
