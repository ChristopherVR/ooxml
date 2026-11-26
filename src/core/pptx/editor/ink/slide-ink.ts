import { inkBoundsOf, MAX_INK_SAMPLES } from '../../../drawingml/ink-geometry';
import type { InkCoordinate } from '../../../drawingml/ink-geometry';
import { recognizeInkShape } from '../../../drawingml/ink-recognition';
import { pointInPolygon } from '../../../geometry/shape-boolean-types';
import type { ContentPartInkStroke, PptxElement } from '../../core/types';
import { canInteractWithElement } from '../render/element-locks';

export interface SlideInkStroke {
	owner: PptxElement;
	index: number;
	points: InkCoordinate[];
	style: Pick<
		ContentPartInkStroke,
		'color' | 'width' | 'opacity' | 'pointTimestamps' | 'startTimestamp' | 'durationMs'
	>;
}

const MAX_SLIDE_SAMPLES = 100_000;
const MAX_SLIDE_STROKES = 1_000;

function slidePoints(element: PptxElement, path: string): InkCoordinate[] {
	if (path.length > MAX_INK_SAMPLES * 64) {
		return [];
	}
	const samples = path.matchAll(
		/([ML])\s*(-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)[,\s]+(-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)/giu,
	);
	const angle = ((element.rotation ?? 0) * Math.PI) / 180;
	const cos = Math.cos(angle);
	const sin = Math.sin(angle);
	const cx = element.width / 2;
	const cy = element.height / 2;
	const points: InkCoordinate[] = [];
	let cursor = 0;
	for (const match of samples) {
		const command = match[1];
		if (
			command !== (points.length ? 'L' : 'M') ||
			path.slice(cursor, match.index).trim() ||
			points.length >= MAX_INK_SAMPLES
		) {
			return [];
		}
		cursor = match.index + match[0].length;
		const x = (Number(match[2]) - cx) * (element.flipHorizontal ? -1 : 1);
		const y = (Number(match[3]) - cy) * (element.flipVertical ? -1 : 1);
		const px = element.x + cx + x * cos - y * sin;
		const py = element.y + cy + x * sin + y * cos;
		if (!Number.isFinite(px) || !Number.isFinite(py)) {
			return [];
		}
		points.push({ x: px, y: py });
	}
	return path.slice(cursor).trim() ? [] : points;
}

/** Project top-level ink strokes into the editable slide's coordinate space. */
export function collectSlideInk(elements: readonly PptxElement[]): SlideInkStroke[] {
	const strokes: SlideInkStroke[] = [];
	let total = 0;
	let examined = 0;
	for (const owner of elements) {
		const paths =
			owner.type === 'ink'
				? owner.inkPaths
				: owner.type === 'contentPart'
					? owner.inkStrokes
					: undefined;
		if (!paths) {
			continue;
		}
		for (const [index, stroke] of paths.entries()) {
			if (++examined > MAX_SLIDE_STROKES) return strokes;
			const points = slidePoints(owner, typeof stroke === 'string' ? stroke : stroke.path);
			total += points.length;
			if (total > MAX_SLIDE_SAMPLES || strokes.length >= MAX_SLIDE_STROKES) {
				return strokes;
			}
			if (points.length) {
				const style =
					typeof stroke !== 'string'
						? stroke
						: owner.type === 'ink'
							? {
									color: owner.inkColors?.[index] ?? '#000000',
									width: owner.inkWidths?.[index] ?? 2,
									opacity: owner.inkOpacities?.[index] ?? 1,
									pointTimestamps: owner.inkPointTimestamps?.[index],
								}
							: undefined;
				if (style) {
					strokes.push({ owner, index, points, style });
				}
			}
		}
	}
	return strokes;
}

export function lassoSlideInk(
	strokes: readonly SlideInkStroke[],
	polygon: InkCoordinate[],
): SlideInkStroke[] {
	if (polygon.length < 3 || polygon.length > 256) {
		return [];
	}
	const bounds = inkBoundsOf(polygon);
	return strokes.filter(
		({ points, owner }) =>
			canInteractWithElement(owner, 'select') &&
			points.every(
				(point) =>
					point.x >= bounds.x &&
					point.x <= bounds.x + bounds.width &&
					point.y >= bounds.y &&
					point.y <= bounds.y + bounds.height &&
					pointInPolygon(point, polygon),
			),
	);
}

/** Replace only selected strokes. Other traces in the same part survive. */
export function replaceSlideInk(
	elements: readonly PptxElement[],
	selected: readonly SlideInkStroke[],
	replacement?: PptxElement,
): PptxElement[] {
	const chosen = new Map<PptxElement, Set<number>>();
	for (const { owner, index } of selected) {
		const indices = chosen.get(owner) ?? new Set<number>();
		indices.add(index);
		chosen.set(owner, indices);
	}
	let inserted = false;
	return elements.flatMap((element) => {
		const indices = chosen.get(element);
		if (!indices) {
			return [element];
		}
		const remaining = withoutStrokes(element, indices);
		if (inserted || !replacement) {
			return remaining ? [remaining] : [];
		}
		inserted = true;
		return remaining ? [remaining, replacement] : [replacement];
	});
}

function withoutStrokes(
	element: PptxElement,
	indices: ReadonlySet<number>,
): PptxElement | undefined {
	const remaining = <T>(values: T[]) => values.filter((_, index) => !indices.has(index));
	if (element.type === 'contentPart' && element.inkStrokes) {
		const inkStrokes = remaining(element.inkStrokes);
		return inkStrokes.length ? { ...element, inkStrokes } : undefined;
	}
	if (element.type === 'ink') {
		const inkPaths = remaining(element.inkPaths);
		return inkPaths.length
			? {
					...element,
					inkPaths,
					inkColors: element.inkColors && remaining(element.inkColors),
					inkWidths: element.inkWidths && remaining(element.inkWidths),
					inkOpacities: element.inkOpacities && remaining(element.inkOpacities),
					inkPointPressures: element.inkPointPressures && remaining(element.inkPointPressures),
					inkPointTiltX: element.inkPointTiltX && remaining(element.inkPointTiltX),
					inkPointTiltY: element.inkPointTiltY && remaining(element.inkPointTiltY),
					inkPointTimestamps: element.inkPointTimestamps && remaining(element.inkPointTimestamps),
				}
			: undefined;
	}
	return element;
}

export function convertSlideInk(
	strokes: readonly SlideInkStroke[],
	id: string,
): PptxElement | undefined {
	const first = strokes[0];
	if (
		!first ||
		strokes.length > 3 ||
		strokes.some(({ owner }) => !canInteractWithElement(owner, 'changeShapeType'))
	) {
		return undefined;
	}
	const shape = recognizeInkShape(strokes.map(({ points }) => points));
	const shapeStyle = {
		fillMode: 'none' as const,
		strokeColor: first.style.color,
		strokeWidth: first.style.width,
	};
	if (!shape) {
		return undefined;
	}
	if ('box' in shape) {
		return {
			id,
			type: 'shape',
			shapeType: shape.kind,
			...shape.box,
			rotation: shape.rotation,
			shapeStyle,
		};
	}
	const { start, end } = shape;
	return {
		id,
		type: 'connector',
		shapeType: 'line',
		...inkBoundsOf([start, end]),
		flipHorizontal: start.x > end.x,
		flipVertical: start.y > end.y,
		shapeStyle:
			shape.kind === 'arrow' ? { ...shapeStyle, connectorEndArrow: 'triangle' } : shapeStyle,
	};
}
