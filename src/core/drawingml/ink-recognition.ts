import type { InkCoordinate } from './ink-geometry';
import {
	inkBoundsOf,
	inkDistance,
	MAX_INK_SAMPLES,
	projectInkPoint,
	simplifyInk,
} from './ink-geometry';

export type RecognizedInkShape =
	| { kind: 'rect' | 'ellipse' | 'triangle'; box: ReturnType<typeof inkBoundsOf>; rotation: number }
	| { kind: 'line' | 'arrow'; start: InkCoordinate; end: InkCoordinate };

function straight(
	points: readonly InkCoordinate[],
): points is readonly [InkCoordinate, ...InkCoordinate[]] {
	const start = points[0];
	const end = points.at(-1);
	if (!start || !end) {
		return false;
	}
	const length = inkDistance(start, end);
	return (
		length >= 8 &&
		points.every((point) => {
			const projection = projectInkPoint(point, start, end);
			return (
				projection.offset >= -0.04 &&
				projection.offset <= 1.04 &&
				inkDistance(point, projection.point) <= Math.max(2, length * 0.04)
			);
		})
	);
}

function arrowHead(start: InkCoordinate, tip: InkCoordinate, a: InkCoordinate, b: InkCoordinate) {
	const length = inkDistance(start, tip);
	const cross = (point: InkCoordinate) =>
		((tip.x - start.x) * (point.y - tip.y) - (tip.y - start.y) * (point.x - tip.x)) / length;
	return (
		cross(a) * cross(b) < 0 &&
		[a, b].every((point) => {
			const { offset } = projectInkPoint(point, start, tip);
			return offset > 0.48 && offset < 0.98 && Math.abs(cross(point)) > length * 0.045;
		})
	);
}

function recognizeArrow(paths: readonly InkCoordinate[][]): RecognizedInkShape | undefined {
	for (const shaft of paths) {
		const first = shaft[0];
		const last = shaft.at(-1);
		if (!first || !last || !straight(shaft)) {
			continue;
		}
		for (const [start, end] of [
			[first, last],
			[last, first],
		] as const) {
			const head = paths
				.filter((path) => path !== shaft)
				.flatMap((path) => simplifyInk(path, Math.max(2, inkDistance(start, end) * 0.045)));
			const tipRadius = Math.max(4, inkDistance(start, end) * 0.1);
			const tips = head.filter((point) => inkDistance(point, end) <= tipRadius);
			const wings = head.filter((point) => inkDistance(point, end) > tipRadius);
			const [a, b] = wings;
			if (tips.length > 0 && wings.length === 2 && a && b && arrowHead(start, end, a, b)) {
				return { kind: 'arrow', start, end };
			}
		}
	}
	return undefined;
}

function closedShape(points: InkCoordinate[], size: number): RecognizedInkShape | undefined {
	const corners = simplifyInk(points, Math.max(2, size * 0.085)).slice(0, -1);
	const box = inkBoundsOf(points);
	if (corners.length === 3 && box.width > 5 && box.height > 5) {
		// DrawingML's triangle has an apex at the top. Rotate the preset to the
		// observed base, using the longest edge when the tip is ambiguous.
		const edges = corners.flatMap((a, index) => {
			const b = corners[(index + 1) % 3];
			const apex = corners[(index + 2) % 3];
			return b && apex ? [{ a, b, apex }] : [];
		});
		const base = edges.sort((a, b) => inkDistance(b.a, b.b) - inkDistance(a.a, a.b))[0];
		if (base) {
			const { a, b, apex } = base;
			const width = inkDistance(a, b);
			const projection = projectInkPoint(apex, a, b);
			const height = inkDistance(apex, projection.point);
			if (width > 0 && height / width > 0.2 && Math.abs(projection.offset - 0.5) < 0.26) {
				const midX = (a.x + b.x) / 2;
				const midY = (a.y + b.y) / 2;
				return {
					kind: 'triangle',
					box: { x: (midX + apex.x - width) / 2, y: (midY + apex.y - height) / 2, width, height },
					rotation:
						(Math.atan2(apex.x - projection.point.x, projection.point.y - apex.y) * 180) / Math.PI,
				};
			}
		}
	}
	const [a, b, c, d] = corners;
	if (corners.length === 4 && a && b && c && d) {
		const angle = Math.atan2(b.y - a.y, b.x - a.x);
		const cos = Math.cos(angle);
		const sin = Math.sin(angle);
		const local = (point: InkCoordinate) => ({
			x: point.x * cos + point.y * sin,
			y: point.y * cos - point.x * sin,
		});
		const [p, q, r, s] = [local(a), local(b), local(c), local(d)];
		const tolerance = size * 0.16;
		if (
			Math.abs(q.x - r.x) < tolerance &&
			Math.abs(r.y - s.y) < tolerance &&
			Math.abs(s.x - p.x) < tolerance
		) {
			const fit = inkBoundsOf(points.map(local));
			const cx = fit.x + fit.width / 2;
			const cy = fit.y + fit.height / 2;
			if (Math.min(fit.width, fit.height) > size * 0.15) {
				return {
					kind: 'rect',
					box: {
						x: cx * cos - cy * sin - fit.width / 2,
						y: cx * sin + cy * cos - fit.height / 2,
						width: fit.width,
						height: fit.height,
					},
					rotation: (angle * 180) / Math.PI,
				};
			}
		}
	}
	const rx = box.width / 2;
	const ry = box.height / 2;
	if (points.length >= 9 && Math.min(rx, ry) >= 4) {
		const onEllipse = points.filter(
			(point) =>
				Math.abs(1 - Math.hypot((point.x - box.x - rx) / rx, (point.y - box.y - ry) / ry)) <= 0.18,
		);
		if (onEllipse.length / points.length >= 0.95 && corners.length >= 5) {
			return { kind: 'ellipse', box, rotation: 0 };
		}
	}
	return undefined;
}

export function recognizeInkShape(
	paths: readonly InkCoordinate[][],
): RecognizedInkShape | undefined {
	if (
		!paths.length ||
		paths.length > 3 ||
		paths.some(
			(path) =>
				path.length < 2 ||
				path.length > MAX_INK_SAMPLES ||
				path.some(({ x, y }) => !Number.isFinite(x) || !Number.isFinite(y)),
		)
	) {
		return undefined;
	}
	if (paths.length > 1) {
		return recognizeArrow(paths);
	}
	const points = paths[0];
	const start = points?.[0];
	const end = points?.at(-1);
	if (!points || !start || !end) {
		return undefined;
	}
	if (straight(points)) {
		return { kind: 'line', start, end };
	}
	const { width, height } = inkBoundsOf(points);
	const size = Math.hypot(width, height);
	if (size < 12) {
		return undefined;
	}
	const [a, tip, wing, returnTip, last, ...extra] = simplifyInk(points, size * 0.045);
	if (
		a &&
		tip &&
		wing &&
		returnTip &&
		last &&
		extra.length === 0 &&
		inkDistance(tip, returnTip) < size * 0.12 &&
		arrowHead(a, tip, wing, last)
	) {
		return { kind: 'arrow', start: a, end: tip };
	}
	if (inkDistance(start, end) <= size * 0.15) {
		return closedShape([...points.slice(0, -1), start], size);
	}
	return undefined;
}
