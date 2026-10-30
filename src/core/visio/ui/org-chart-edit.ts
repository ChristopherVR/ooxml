import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioShapePageBounds } from './diagram-parts';
import { layoutOrgPage } from './org-chart-layout';
import { ORG_STENCIL_MASTER_IDS, orgStencilBytes } from './org-chart-stencil-data';
import {
	visioIsOrgAssistant,
	visioOrgDataEdits,
	visioOrgHierarchy,
	visioOrgText,
	type VisioOrgPositionData,
} from './org-chart';

/** The org chart masters of Visio's Organization Chart stencil that a position can be. */
export type VisioOrgMaster = 'Executive Belt' | 'Manager Belt' | 'Position Belt' | 'Assistant Belt';
export type VisioOrgAddKind = 'subordinate' | 'peer' | 'assistant';

/** Every Belt master is this large, in drawing inches (the master's User.Width and Height). */
export const VISIO_ORG_SIZE = { width: 1.625, height: 0.875 } as const;
/** Shape IDs an instance of a Belt master takes: the group and its nine sub-shapes, with room. */
const ID_SPAN = 16;

interface Box {
	minX: number;
	maxX: number;
	minY: number;
	maxY: number;
}
/** A shape's bounds in drawing inches, y up. */
function drawingBox(page: VisioPage, shape: VisioShape): Box {
	const box = visioShapePageBounds(shape);
	const ratio = page.drawingToPageScale ?? 1;
	return {
		minX: box.minX / ratio,
		maxX: box.maxX / ratio,
		minY: box.minY / ratio,
		maxY: box.maxY / ratio,
	};
}
/** The largest numeric shape ID in the page tree. */
function maxShapeId(page: VisioPage): number {
	let max = 0;
	const pending = [...page.shapes];
	while (pending.length) {
		const shape = pending.pop()!;
		if (/^\d{1,10}$/.test(shape.id)) max = Math.max(max, Number(shape.id));
		pending.push(...shape.children);
	}
	return max;
}
/** IDs for `positions` new positions (each leaves room for its sub-shapes), then `lines` connectors. */
export function visioOrgFreeIds(
	page: VisioPage,
	positions: number,
	lines: number,
): { positions: string[]; lines: string[] } | undefined {
	const max = maxShapeId(page);
	const last = max + (positions + lines) * ID_SPAN;
	if (last >= 4294967295) return undefined;
	const at = (index: number) => String(max + 1 + index * ID_SPAN);
	return {
		positions: Array.from({ length: positions }, (_, index) => at(index)),
		lines: Array.from({ length: lines }, (_, index) => at(positions + index)),
	};
}

/** The drop of one position: an instance of the Belt master `master`, centred on `at`. */
export function visioOrgDropPosition(
	pageId: string,
	shapeId: string,
	master: VisioOrgMaster,
	at: { x: number; y: number },
): VisioEdit {
	return {
		type: 'drop-stencil-master',
		pageId,
		shapeId,
		master: ORG_STENCIL_MASTER_IDS[master],
		stencil: orgStencilBytes(),
		x: at.x,
		y: at.y,
	};
}

/**
 * A reporting line as Visio draws one: an instance of the Organization Chart stencil's Dynamic
 * connector, dropped between the two positions and then glued to both, so it follows them. The
 * drop edit comes first in the returned pair (drops run before the rest of a batch); `glue`
 * follows the edits that place the positions.
 */
export function visioOrgReportingLine(
	pageId: string,
	lineId: string,
	manager: string,
	subordinate: string,
	between: { x: number; y: number },
): { drop: VisioEdit; glue: VisioEdit[] } {
	return {
		drop: {
			type: 'drop-stencil-master',
			pageId,
			shapeId: lineId,
			master: ORG_STENCIL_MASTER_IDS['Dynamic connector'],
			stencil: orgStencilBytes(),
			x: between.x,
			y: between.y,
		},
		glue: [
			{ type: 'glue-connector', pageId, shapeId: lineId, endpoint: 'begin', target: manager },
			{ type: 'glue-connector', pageId, shapeId: lineId, endpoint: 'end', target: subordinate },
		],
	};
}

const masterOf = (shape: VisioShape): VisioOrgMaster => {
	if (visioIsOrgAssistant(shape)) return 'Assistant Belt';
	if (/^Executive Belt/i.test(shape.name)) return 'Executive Belt';
	return /^Manager Belt/i.test(shape.name) ? 'Manager Belt' : 'Position Belt';
};

/**
 * Org Chart > Subordinate, Assistant or Coworker: a new position under the selected one (or under
 * its manager, for a coworker; a coworker of an assistant is another assistant), an instance of
 * Visio's Position, Assistant or the selected position's Belt master, with a Dynamic connector
 * glued to it and its manager, then the whole tree laid out again so the new box has room. An
 * assistant sits beside its manager's line. One edit batch, so it is one undo step. The new
 * shape id is returned for selection.
 */
export function visioOrgAddPosition(
	page: VisioPage,
	shapeId: string,
	kind: VisioOrgAddKind,
): { edits: VisioEdit[]; newId: string } | undefined {
	const hierarchy = visioOrgHierarchy(page);
	const selected = hierarchy.positions.find((shape) => shape.id === shapeId);
	if (!selected) return undefined;
	const managerId = kind === 'peer' ? hierarchy.parentOf(shapeId) : shapeId;
	const manager = hierarchy.positions.find((shape) => shape.id === managerId);
	const ids = visioOrgFreeIds(page, 1, 1);
	if (!manager || !ids) return undefined;
	const newId = ids.positions[0]!;
	const master: VisioOrgMaster =
		kind === 'assistant'
			? 'Assistant Belt'
			: kind === 'peer'
				? masterOf(selected)
				: 'Position Belt';
	const assistant = master === 'Assistant Belt';
	const planned = layoutOrgPage(page, [
		{ id: newId, ...VISIO_ORG_SIZE, parent: manager.id, ...(assistant ? { assistant } : {}) },
	]);
	const centre = planned?.centres.get(newId);
	if (!planned || !centre) return undefined;
	const data: VisioOrgPositionData = { Name: 'Name', Title: assistant ? 'Assistant' : 'Position' };
	const managerBox = drawingBox(page, manager);
	const line = visioOrgReportingLine(page.id, ids.lines[0]!, manager.id, newId, {
		x: (managerBox.minX + managerBox.maxX + centre.x * 2) / 4,
		y: (managerBox.minY + centre.y) / 2,
	});
	const pageId = page.id;
	return {
		newId,
		edits: [
			visioOrgDropPosition(pageId, newId, master, centre),
			line.drop,
			{ type: 'replace-plain-text', pageId, shapeId: newId, text: visioOrgText(data) },
			...visioOrgDataEdits(pageId, newId, data),
			...planned.moves,
			...line.glue,
		],
	};
}

/**
 * Org Chart > Delete Position: the position and the connectors glued to it go; its subordinates
 * (and assistants) are joined to its manager, or become tops of their own trees when it had none.
 * One batch, so the whole command is one undo step.
 */
export function visioOrgDeletePosition(page: VisioPage, shapeId: string): VisioEdit[] | undefined {
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
	const ids = manager ? visioOrgFreeIds(page, 0, children.length) : { positions: [], lines: [] };
	if (!ids) return undefined;
	const joins = manager
		? children.map((child, index) => {
				const a = drawingBox(page, manager);
				const b = drawingBox(page, child);
				return visioOrgReportingLine(page.id, ids.lines[index]!, manager.id, child.id, {
					x: (a.minX + a.maxX + b.minX + b.maxX) / 4,
					y: (a.minY + b.maxY) / 2,
				});
			})
		: [];
	const removal = [...connectors, shapeId].map((id): VisioEdit => ({
		type: 'delete-shape',
		pageId: page.id,
		shapeId: id,
	}));
	// The new lines are glued before anything is deleted: the deletion changes the page the glue
	// would be worked out on.
	return [...joins.map((join) => join.drop), ...joins.flatMap((join) => join.glue), ...removal];
}
