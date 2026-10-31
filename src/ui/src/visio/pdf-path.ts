import type { PathSegment } from 'ooxml-core/visio';

export { MAX_PATH_CHARS, MAX_PATH_SEGMENTS, parsePath, type PathSegment } from 'ooxml-core/visio';

const n = (value: number) => (Number.isFinite(value) ? String(Math.round(value * 1e6) / 1e6) : '0');

/** PDF path-construction operators for the segments. */
export function pathOperators(segments: readonly PathSegment[]): string[] {
	return segments.map((segment) => {
		switch (segment.t) {
			case 'M':
				return `${n(segment.x)} ${n(segment.y)} m`;
			case 'L':
				return `${n(segment.x)} ${n(segment.y)} l`;
			case 'C':
				return `${n(segment.x1)} ${n(segment.y1)} ${n(segment.x2)} ${n(segment.y2)} ${n(segment.x)} ${n(segment.y)} c`;
			default:
				return 'h';
		}
	});
}

export interface PathEnd {
	x: number;
	y: number;
	/** Direction of travel at the end, in radians; for the start it points along the path. */
	angle: number;
}

/** Position and tangent where the first and last drawn segments meet the path ends. */
export function pathEnds(segments: readonly PathSegment[]): { start?: PathEnd; end?: PathEnd } {
	const points: { x: number; y: number }[][] = [];
	let sub: { x: number; y: number }[] = [];
	let first: { x: number; y: number } | undefined;
	for (const segment of segments) {
		if (segment.t === 'M') {
			sub = [{ x: segment.x, y: segment.y }];
			first = sub[0];
			points.push(sub);
		} else if (segment.t === 'Z') {
			if (first) sub.push(first);
		} else if (segment.t === 'L') sub.push({ x: segment.x, y: segment.y });
		else {
			sub.push({ x: segment.x1, y: segment.y1 }, { x: segment.x2, y: segment.y2 });
			sub.push({ x: segment.x, y: segment.y });
		}
	}
	const head = points[0],
		tail = points[points.length - 1];
	if (!head || !tail || head.length < 2 || tail.length < 2) return {};
	const direction = (
		from: { x: number; y: number },
		list: { x: number; y: number }[],
		step: 1 | -1,
	) => {
		let index = list.indexOf(from) + step;
		while (list[index]) {
			const p = list[index]!;
			if (Math.hypot(p.x - from.x, p.y - from.y) > 1e-9)
				return Math.atan2(p.y - from.y, p.x - from.x);
			index += step;
		}
		return 0;
	};
	const a = head[0]!,
		b = tail[tail.length - 1]!;
	return {
		start: { x: a.x, y: a.y, angle: direction(a, head, 1) },
		end: { x: b.x, y: b.y, angle: direction(b, tail, -1) + Math.PI },
	};
}
