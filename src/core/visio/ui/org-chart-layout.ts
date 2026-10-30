import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioShapePageBounds } from './diagram-parts';
import { visioIsOrgAssistant, visioOrgHierarchy } from './org-chart';
import { visioPageEditToDrawing } from './page-edit';

/** Page inches between neighbours of a level, and between levels. */
const SPACING = 0.4;
const LAYER_SPACING = 0.6;
/** Where an assistant sits beside its manager's line, and the space between stacked assistants. */
const ASSISTANT_GAP = 0.4;
const ASSISTANT_STACK = 0.15;

export interface OrgLayoutNode {
	id: string;
	/** Page inches. */
	width: number;
	height: number;
	/** An assistant sits beside its manager's line, between the manager and the subordinates. */
	assistant?: boolean;
	/** Where the position is now (x, page inches): siblings keep their order. New ones go last. */
	hint?: number;
}
export interface OrgPoint {
	x: number;
	y: number;
}

interface Placed {
	node: OrgLayoutNode;
	depth: number;
	children: Placed[];
	assistants: OrgLayoutNode[];
	/** Extents of the whole subtree left and right of the node's centre. */
	left: number;
	right: number;
	/** Centres of the children relative to this node's centre. */
	offsets: number[];
}

/**
 * Lay a tree of positions out top-down and return each position's centre, x rightward and y
 * downward from the top left of the whole chart. Each manager sits over its subordinates, and
 * every level is a row. A manager's assistants are stacked to the right of the line that drops
 * from the manager, in a band of their own between the manager's row and the next, as Visio
 * places them; the rows below are pushed down to make room.
 */
export function planOrgLayout(
	nodes: readonly OrgLayoutNode[],
	edges: readonly { from: string; to: string }[],
): Map<string, OrgPoint> | undefined {
	const byId = new Map(nodes.map((node) => [node.id, node]));
	const parentOf = new Map<string, string>();
	for (const edge of edges)
		if (byId.has(edge.from) && byId.has(edge.to) && edge.from !== edge.to && !parentOf.has(edge.to))
			parentOf.set(edge.to, edge.from);
	// A cycle has no top: its members are laid out as tops of their own trees.
	for (const node of nodes) {
		const seen = new Set<string>([node.id]);
		for (let up = parentOf.get(node.id); up !== undefined; up = parentOf.get(up)) {
			if (seen.has(up)) {
				parentOf.delete(node.id);
				break;
			}
			seen.add(up);
		}
	}
	const isAssistant = (node: OrgLayoutNode) =>
		node.assistant === true &&
		parentOf.has(node.id) &&
		!byId.get(parentOf.get(node.id)!)!.assistant;
	const assistantsOf = new Map<string, OrgLayoutNode[]>();
	const childrenOf = new Map<string, OrgLayoutNode[]>();
	for (const node of nodes) {
		const parent = parentOf.get(node.id);
		if (parent === undefined) continue;
		const into = isAssistant(node) ? assistantsOf : childrenOf;
		into.set(parent, [...(into.get(parent) ?? []), node]);
	}
	const byHint = (a: OrgLayoutNode, b: OrgLayoutNode) =>
		(a.hint ?? Infinity) === (b.hint ?? Infinity)
			? 0
			: (a.hint ?? Infinity) < (b.hint ?? Infinity)
				? -1
				: 1;
	const build = (node: OrgLayoutNode, depth: number): Placed => {
		const assistants = assistantsOf.get(node.id) ?? [];
		const kids = [...(childrenOf.get(node.id) ?? [])]
			.sort(byHint)
			.map((kid) => build(kid, depth + 1));
		// Children side by side by their own extents, then centred under this node.
		const centres: number[] = [];
		let cursor = 0;
		kids.forEach((kid, index) => {
			cursor += index ? kids[index - 1]!.right + SPACING + kid.left : 0;
			centres.push(cursor);
		});
		const middle = kids.length ? (centres[0]! + centres.at(-1)!) / 2 : 0;
		const offsets = centres.map((centre) => centre - middle);
		const half = node.width / 2;
		const reach = assistants.length
			? ASSISTANT_GAP + Math.max(...assistants.map((item) => item.width))
			: 0;
		const left = Math.max(half, ...kids.map((kid, index) => kid.left - offsets[index]!));
		const right = Math.max(half, reach, ...kids.map((kid, index) => kid.right + offsets[index]!));
		return { node, depth, children: kids, assistants, left, right, offsets };
	};
	const roots = nodes.filter((node) => !parentOf.has(node.id)).sort(byHint);
	if (!roots.length) return undefined;
	const trees = roots.map((root) => build(root, 0));
	// Row heights, and the band each level reserves for its managers' assistants.
	const rowHeight: number[] = [];
	const band: number[] = [];
	const visit = (item: Placed, action: (value: Placed) => void) => {
		action(item);
		item.children.forEach((kid) => visit(kid, action));
	};
	const bandHeight = (item: Placed) =>
		item.assistants.reduce((sum, value) => sum + value.height, 0) +
		(item.assistants.length - 1) * ASSISTANT_STACK;
	for (const tree of trees)
		visit(tree, (item) => {
			rowHeight[item.depth] = Math.max(rowHeight[item.depth] ?? 0, item.node.height);
			if (item.assistants.length)
				band[item.depth] = Math.max(band[item.depth] ?? 0, bandHeight(item) + LAYER_SPACING);
		});
	const rowTop: number[] = [];
	for (let depth = 0; depth < rowHeight.length; depth++)
		rowTop[depth] = depth
			? rowTop[depth - 1]! + rowHeight[depth - 1]! + LAYER_SPACING + (band[depth - 1] ?? 0)
			: 0;
	const result = new Map<string, OrgPoint>();
	let cursor = 0;
	const place = (item: Placed, centre: number) => {
		const y = rowTop[item.depth]! + rowHeight[item.depth]! / 2;
		result.set(item.node.id, { x: centre, y });
		let top = rowTop[item.depth]! + rowHeight[item.depth]! + LAYER_SPACING;
		for (const value of item.assistants) {
			result.set(value.id, {
				x: centre + ASSISTANT_GAP + value.width / 2,
				y: top + value.height / 2,
			});
			top += value.height + ASSISTANT_STACK;
		}
		item.children.forEach((kid, index) => place(kid, centre + item.offsets[index]!));
	};
	trees.forEach((tree, index) => {
		cursor += index ? trees[index - 1]!.right + SPACING + tree.left : tree.left;
		place(tree, cursor);
	});
	// Origin at the top left of the whole chart.
	const edge = Math.min(...[...result].map(([id, at]) => at.x - byId.get(id)!.width / 2));
	for (const at of result.values()) at.x -= edge;
	return result;
}

export interface OrgNewNode extends OrgLayoutNode {
	/** The position the new one reports to. */
	parent: string;
}
const scale = (page: VisioPage) => {
	const ratio = page.drawingToPageScale ?? 1;
	return Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
};

/**
 * Lay the page's tree out again with `extra` positions (not in the page yet) added. Existing
 * positions keep the top-left corner of their bounds. Returns the move edits for existing
 * positions and the new positions' centres (drawing inches, y up).
 */
export function layoutOrgPage(page: VisioPage, extra: readonly OrgNewNode[] = []) {
	const { positions, lines } = visioOrgHierarchy(page);
	if (!positions.length) return undefined;
	const boxes = new Map<string, ReturnType<typeof visioShapePageBounds>>(
		positions.map((shape: VisioShape) => [shape.id, visioShapePageBounds(shape)]),
	);
	const ratio = scale(page);
	const nodes: OrgLayoutNode[] = [
		...positions.map((shape) => {
			const box = boxes.get(shape.id)!;
			return {
				id: shape.id,
				width: box.maxX - box.minX,
				height: box.maxY - box.minY,
				assistant: visioIsOrgAssistant(shape),
				hint: (box.minX + box.maxX) / 2,
			};
		}),
		...extra.map((item) => ({
			id: item.id,
			width: item.width * ratio,
			height: item.height * ratio,
			...(item.assistant ? { assistant: true } : {}),
		})),
	];
	const placed = planOrgLayout(nodes, [
		...lines,
		...extra.map((item) => ({ from: item.parent, to: item.id })),
	]);
	if (!placed) return undefined;
	const left = Math.min(...[...boxes.values()].map((box) => box.minX));
	const top = Math.max(...[...boxes.values()].map((box) => box.maxY));
	const moves: VisioEdit[] = [];
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
	const centres = new Map<string, OrgPoint>();
	for (const item of extra) {
		const at = placed.get(item.id);
		if (at) centres.set(item.id, { x: (left + at.x) / ratio, y: (top - at.y) / ratio });
	}
	return { moves, centres };
}

/** Org Chart > Layout: put every position back in a tidy top-down tree. */
export function visioOrgReLayout(page: VisioPage): VisioEdit[] | undefined {
	const result = layoutOrgPage(page);
	return result && result.moves.length ? result.moves : undefined;
}
