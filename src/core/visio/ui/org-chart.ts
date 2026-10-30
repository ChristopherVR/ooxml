import { layoutGraph } from '../../geometry/graph-layout';
import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioShapePageBounds } from './diagram-parts';
import { visioConnectorEdges } from './auto-layout';
import { visioPageEditToDrawing } from './page-edit';

/**
 * The Shape Data rows of an org chart position, in the order and with the row names Visio's own
 * org chart masters (Executive, Manager, Position and Assistant Belt) carry, so a position keeps
 * its data when Visio opens the file. The Email row is labelled E-Mail, as Visio does.
 */
export const VISIO_ORG_ROWS = [
	{ row: 'Department', label: 'Department' },
	{ row: 'Telephone', label: 'Telephone' },
	{ row: 'Name', label: 'Name' },
	{ row: 'Title', label: 'Title' },
	{ row: 'Email', label: 'E-Mail' },
] as const;

export interface VisioOrgPositionData {
	Name: string;
	Title: string;
	Department?: string;
	Telephone?: string;
	Email?: string;
}

/** The set-shape-data edits that give a shape the org chart rows. */
export function visioOrgDataEdits(
	pageId: string,
	shapeId: string,
	data: VisioOrgPositionData,
): VisioEdit[] {
	return VISIO_ORG_ROWS.map(({ row, label }) => ({
		type: 'set-shape-data',
		pageId,
		shapeId,
		row,
		data: { label, prompt: '', type: 'string', format: '', value: data[row] ?? '' },
	}));
}

/** Whether a shape is an org chart position: it carries every org chart Shape Data row. */
export const visioIsOrgPosition = (shape: VisioShape | undefined): boolean =>
	!!shape && VISIO_ORG_ROWS.every(({ row }) => shape.shapeData?.some((item) => item.name === row));

/** The org chart position a selected shape is, when it is one. */
export function visioOrgPosition(page: VisioPage, shapeId: string): VisioShape | undefined {
	const shape = page.shapes.find((item) => item.id === shapeId);
	return visioIsOrgPosition(shape) ? shape : undefined;
}

export interface VisioOrgHierarchy {
	/** Every position on the page. */
	positions: VisioShape[];
	/** Reporting lines between positions, manager to subordinate. */
	lines: { from: string; to: string }[];
	parentOf(id: string): string | undefined;
	childrenOf(id: string): string[];
}

/** The positions of a page and the glued connectors that join them. */
export function visioOrgHierarchy(page: VisioPage): VisioOrgHierarchy {
	const positions = page.shapes.filter(visioIsOrgPosition);
	const ids = new Set(positions.map((shape) => shape.id));
	const lines = visioConnectorEdges(page).filter((edge) => ids.has(edge.from) && ids.has(edge.to));
	return {
		positions,
		lines,
		parentOf: (id) => lines.find((line) => line.to === id)?.from,
		childrenOf: (id) => lines.filter((line) => line.from === id).map((line) => line.to),
	};
}

const SIZE = { width: 1.8, height: 0.8 };
const LINE = '#326765';
const FILLS = { subordinate: '#D8EFEB', peer: '#D8EFEB', assistant: '#F3EEE1' } as const;
export type VisioOrgAddKind = keyof typeof FILLS;

interface Box {
	minX: number;
	maxX: number;
	minY: number;
	maxY: number;
}
const scale = (page: VisioPage) => {
	const ratio = page.drawingToPageScale ?? 1;
	return Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
};
/** A shape's bounds in drawing inches, y up. */
function drawingBox(page: VisioPage, shape: VisioShape): Box {
	const box = visioShapePageBounds(shape);
	const ratio = scale(page);
	return {
		minX: box.minX / ratio,
		maxX: box.maxX / ratio,
		minY: box.minY / ratio,
		maxY: box.maxY / ratio,
	};
}
/** Consecutive free shape IDs after the largest ID in the page tree. */
function freeIds(page: VisioPage, count: number): string[] | undefined {
	let max = 0;
	const pending = [...page.shapes];
	while (pending.length) {
		const shape = pending.pop()!;
		if (/^\d{1,10}$/.test(shape.id)) max = Math.max(max, Number(shape.id));
		pending.push(...shape.children);
	}
	if (max + count >= 4294967295) return undefined;
	return Array.from({ length: count }, (_, index) => String(max + 1 + index));
}

interface NewNode {
	id: string;
	width: number;
	height: number;
	parent: string;
}
/**
 * Lay the tree out top-down with the geometry area's hierarchy layout. Existing positions keep
 * the top-left corner of their bounds; `extra` are positions not yet in the page. Returns the
 * move edits for existing positions and the new positions' centres (drawing inches).
 */
function layoutTree(page: VisioPage, extra: NewNode[] = []) {
	const { positions, lines } = visioOrgHierarchy(page);
	if (!positions.length) return undefined;
	const boxes = new Map(positions.map((shape) => [shape.id, visioShapePageBounds(shape)]));
	const ratio = scale(page);
	const nodes = [
		...positions.map((shape) => {
			const box = boxes.get(shape.id)!;
			return { id: shape.id, width: box.maxX - box.minX, height: box.maxY - box.minY };
		}),
		...extra.map((item) => ({ ...item, width: item.width * ratio, height: item.height * ratio })),
	];
	const edges = [...lines, ...extra.map((item) => ({ from: item.parent, to: item.id }))];
	const left = Math.min(...[...boxes.values()].map((box) => box.minX));
	const top = Math.max(...[...boxes.values()].map((box) => box.maxY));
	let placed: ReturnType<typeof layoutGraph>;
	try {
		placed = layoutGraph(nodes, edges, 'hierarchy', { spacing: 0.4, layerSpacing: 0.6 });
	} catch {
		return undefined;
	}
	const moves: VisioEdit[] = [];
	const centres = new Map<string, { x: number; y: number }>();
	for (const shape of positions) {
		const at = placed.get(shape.id);
		const box = boxes.get(shape.id)!;
		const pin = shape.rotation;
		if (!at || !pin) continue;
		const dx = left + at.x - (box.minX + box.maxX) / 2;
		const dy = top - at.y - (box.minY + box.maxY) / 2;
		if (Math.abs(dx) < 1e-9 && Math.abs(dy) < 1e-9) continue;
		moves.push(
			visioPageEditToDrawing(page, {
				type: 'move-shape',
				pageId: page.id,
				shapeId: shape.id,
				x: pin.pinX + dx,
				y: pin.pinY + dy,
			}),
		);
	}
	for (const item of extra) {
		const at = placed.get(item.id);
		if (at) centres.set(item.id, { x: (left + at.x) / ratio, y: (top - at.y) / ratio });
	}
	return { moves, centres };
}

/** Org Chart > Layout: put every position back in a tidy top-down tree. */
export function visioOrgReLayout(page: VisioPage): VisioEdit[] | undefined {
	const result = layoutTree(page);
	return result && result.moves.length ? result.moves : undefined;
}

/** A glued connector from the bottom of a manager's box to the top of a subordinate's. */
const reportingLine = (
	pageId: string,
	id: string,
	from: [string, Box],
	to: [string, Box],
): VisioEdit => ({
	type: 'create-line',
	pageId,
	shapeId: id,
	beginX: (from[1].minX + from[1].maxX) / 2,
	beginY: from[1].minY,
	endX: (to[1].minX + to[1].maxX) / 2,
	endY: to[1].maxY,
	connect: { begin: from[0], end: to[0] },
});

/**
 * Org Chart > Subordinate, Assistant or Coworker: a new position under the selected one (or under
 * its manager, for a coworker), joined by a glued connector and carrying the org chart Shape Data,
 * then the whole tree laid out again so the new box has room. One edit batch, so it is one undo
 * step. The new shape id is returned for selection. Assistants are placed like subordinates.
 */
export function visioOrgAddPosition(
	page: VisioPage,
	shapeId: string,
	kind: VisioOrgAddKind,
): { edits: VisioEdit[]; newId: string } | undefined {
	const hierarchy = visioOrgHierarchy(page);
	if (!hierarchy.positions.some((shape) => shape.id === shapeId)) return undefined;
	const managerId = kind === 'peer' ? hierarchy.parentOf(shapeId) : shapeId;
	const manager = hierarchy.positions.find((shape) => shape.id === managerId);
	const ids = freeIds(page, 2);
	if (!manager || !ids) return undefined;
	const [newId, lineId] = ids as [string, string];
	const planned = layoutTree(page, [{ id: newId, ...SIZE, parent: manager.id }]);
	const centre = planned?.centres.get(newId);
	if (!planned || !centre) return undefined;
	const title = kind === 'assistant' ? 'Assistant' : 'Position';
	const box: Box = {
		minX: centre.x - SIZE.width / 2,
		maxX: centre.x + SIZE.width / 2,
		minY: centre.y - SIZE.height / 2,
		maxY: centre.y + SIZE.height / 2,
	};
	const pageId = page.id;
	return {
		newId,
		edits: [
			{ type: 'drop-stencil-master', pageId, shapeId: newId, master: 'rectangle', ...centre },
			{ type: 'resize-shape', pageId, shapeId: newId, ...SIZE },
			{ type: 'replace-plain-text', pageId, shapeId: newId, text: `${title}\nName` },
			{ type: 'format-shape', pageId, shapeId: newId, fillColor: FILLS[kind], lineColor: LINE },
			...visioOrgDataEdits(pageId, newId, { Name: 'Name', Title: title }),
			...planned.moves,
			reportingLine(pageId, lineId, [manager.id, drawingBox(page, manager)], [newId, box]),
		],
	};
}

/**
 * Org Chart > Delete Position: the position and the connectors glued to it go; its subordinates
 * are joined to its manager, or become tops of their own trees when it had none. Two batches
 * (the deletion, then the new connectors), applied in order.
 */
export function visioOrgDeletePosition(
	page: VisioPage,
	shapeId: string,
): VisioEdit[][] | undefined {
	const hierarchy = visioOrgHierarchy(page);
	if (!hierarchy.positions.some((shape) => shape.id === shapeId)) return undefined;
	const connectors = new Set(
		page.connectors
			.filter((connection) => connection.toShapeId === shapeId)
			.map((connection) => connection.fromShapeId),
	);
	const manager = hierarchy.positions.find((shape) => shape.id === hierarchy.parentOf(shapeId));
	const children = hierarchy.positions.filter((shape) =>
		hierarchy.childrenOf(shapeId).includes(shape.id),
	);
	const ids = manager ? freeIds(page, children.length) : [];
	if (!ids) return undefined;
	// Stencil shapes are deleted in a transaction of their own, before the new connectors go in.
	const removal = [...connectors, shapeId].map((id): VisioEdit => ({
		type: 'delete-shape',
		pageId: page.id,
		shapeId: id,
	}));
	const joins = manager
		? children.map((child, index) =>
				reportingLine(
					page.id,
					ids[index]!,
					[manager.id, drawingBox(page, manager)],
					[child.id, drawingBox(page, child)],
				),
			)
		: [];
	return joins.length ? [removal, joins] : [removal];
}
