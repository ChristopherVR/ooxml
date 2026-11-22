import { svgPathToSegments } from '../../geometry/svg-path-segments';
import type { VisioPathCreateEdit, VisioPathStroke } from '../edit-path-commands';
import type { VisioPage } from '../model';
import { visioNextShapeId } from './shape-id';

interface VisioIcon {
	name: string;
	path: string;
	color: string;
}

/** Converts the gallery's 24-by-24 stroked SVG paths to native, editable Visio geometry. */
export function visioIconInsertCommands(
	page: VisioPage,
	icons: readonly VisioIcon[],
): VisioPathCreateEdit[] {
	if (!icons.length) return [];
	const ratio = page.drawingToPageScale ?? 1;
	const columns = Math.ceil(Math.sqrt(icons.length));
	const rows = Math.ceil(icons.length / columns);
	const pitch = Math.min(1.2, page.width / (columns + 1), page.height / (rows + 1)) / ratio;
	const size = pitch * 0.8;
	const firstId = Number(visioNextShapeId(page));
	return icons.map((icon, index) => {
		const left = (page.width / ratio - columns * pitch) / 2 + ((index % columns) + 0.1) * pitch;
		const top =
			(page.height / ratio + rows * pitch) / 2 - (Math.floor(index / columns) + 0.1) * pitch;
		const x = (n: number) => left + (n * size) / 24;
		const y = (n: number) => top - (n * size) / 24;
		const paths: VisioPathStroke[] = [];
		let stroke: VisioPathStroke | undefined;
		let current = { x: 0, y: 0 };
		for (const segment of svgPathToSegments(icon.path)) {
			if (segment.type === 'moveTo') {
				current = { x: x(segment.pt.x), y: y(segment.pt.y) };
				stroke = { ...current, segments: [] };
				paths.push(stroke);
			} else if (stroke && segment.type === 'lineTo') {
				current = { x: x(segment.pt.x), y: y(segment.pt.y) };
				stroke.segments.push({ kind: 'line', ...current });
			} else if (stroke && (segment.type === 'cubicBezTo' || segment.type === 'quadBezTo')) {
				const [p1, p2, p3] = segment.pts;
				const control = { x: x(p1.x), y: y(p1.y) };
				const end = p3 ?? p2;
				const next = { x: x(end.x), y: y(end.y) };
				const x1 = p3 ? control.x : (current.x + 2 * control.x) / 3;
				const y1 = p3 ? control.y : (current.y + 2 * control.y) / 3;
				const x2 = p3 ? x(p2.x) : (next.x + 2 * control.x) / 3;
				const y2 = p3 ? y(p2.y) : (next.y + 2 * control.y) / 3;
				stroke.segments.push({ kind: 'cubic', ...next, x1, y1, x2, y2 });
				current = next;
			} else if (stroke && segment.type === 'close') {
				current = { x: stroke.x, y: stroke.y };
				stroke.segments.push({ kind: 'line', ...current });
			}
		}
		const [first, ...subpaths] = paths.filter((path) => path.segments.length > 0);
		if (!first) throw new Error('The icon has no drawable geometry.');
		return {
			...first,
			type: 'create-path',
			pageId: page.id,
			shapeId: String(firstId + index),
			subpaths,
			stroke: { name: icon.name, color: icon.color, weight: (size * 1.75) / 24 },
		};
	});
}
