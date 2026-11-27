import { MAX_INK_SAMPLES, type InkCoordinate } from '../ink-geometry';

const NUMBER = String.raw`-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?`;
const COMMAND = new RegExp(String.raw`([ML])\s*(${NUMBER})[,\s]+(${NUMBER})`, 'giu');

/**
 * The points of a stroke path (`M x y L x y ...`); curves and malformed paths read as none. The
 * result is shared between callers: do not change it.
 */
export function inkPathPoints(path: string): InkCoordinate[] {
	const cached = parsed.get(path);
	if (cached) return cached;
	const points = parse(path);
	// The eraser, lasso and painters read the same strokes again on every move: keep the last few.
	if (parsed.size >= CACHE_SIZE) parsed.delete(parsed.keys().next().value as string);
	parsed.set(path, points);
	return points;
}

const CACHE_SIZE = 2048;
const parsed = new Map<string, InkCoordinate[]>();

function parse(path: string): InkCoordinate[] {
	if (path.length > MAX_INK_SAMPLES * 64) return [];
	const points: InkCoordinate[] = [];
	for (const match of path.matchAll(COMMAND)) {
		const x = Number(match[2]);
		const y = Number(match[3]);
		if (!Number.isFinite(x) || !Number.isFinite(y) || points.length >= MAX_INK_SAMPLES) return [];
		points.push({ x, y });
	}
	return points;
}
