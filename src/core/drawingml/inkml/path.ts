import { MAX_INK_SAMPLES, type InkCoordinate } from '../ink-geometry';

const NUMBER = String.raw`-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?`;
const COMMAND = new RegExp(String.raw`([ML])\s*(${NUMBER})[,\s]+(${NUMBER})`, 'giu');

/** The points of a stroke path (`M x y L x y ...`); curves and malformed paths read as none. */
export function inkPathPoints(path: string): InkCoordinate[] {
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
