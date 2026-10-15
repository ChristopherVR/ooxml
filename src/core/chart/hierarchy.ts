/**
 * The category hierarchy behind the treemap and sunburst families, shared by every format.
 * Moved from the PowerPoint renderer (`chart-treemap-hierarchy.ts`).
 *
 * An Office 2016+ chart (`cx:`) stores its categories leaf first: `cx:strDim type="cat"` holds
 * one `cx:lvl` per level, the first being the leaf labels and the last the outermost parents, each
 * parent label written only on the first row it covers. {@link normalizeLevels} fills those gaps
 * and the builders turn the levels plus the leaf values into a weighted tree.
 */
import { squarify, type TreemapBox } from './treemap-squarify';

export type ParentLabelLayout = 'none' | 'banner' | 'overlapping';

export interface HierarchyNode {
	label: string;
	/** A leaf's absolute value; a parent's is the sum of its children (see {@link aggregate}). */
	weight: number;
	children?: HierarchyNode[];
	seriesIndex?: number;
	pointIndex?: number;
	colorIndex: number;
	parentLabelLayout: ParentLabelLayout;
}

/** One series of a hierarchy: its leaf values and the layout of its parent labels. */
export interface HierarchySeries {
	name: string;
	values: readonly number[];
	parentLabelLayout?: ParentLabelLayout | undefined;
}

/** Fills the blanks of each level with the label above them (parents are written once). */
export function normalizeLevels(
	levels: ReadonlyArray<ReadonlyArray<string>>,
	pointCount: number,
): string[][] {
	return levels.map((level) => {
		let previous = '';
		return Array.from({ length: pointCount }, (_, index) => {
			const label = level[index]?.trim() ?? '';
			if (label) previous = label;
			return label || previous;
		});
	});
}

function findOrAddParent(
	children: HierarchyNode[],
	label: string,
	colorIndex: number,
	parentLabelLayout: ParentLabelLayout,
): HierarchyNode & { children: HierarchyNode[] } {
	const existing = children.find((child) => child.children && child.label === label);
	if (existing?.children) return existing as HierarchyNode & { children: HierarchyNode[] };
	const parent = { label, weight: 0, children: [], colorIndex, parentLabelLayout };
	children.push(parent);
	return parent;
}

function buildSeriesHierarchy(
	series: HierarchySeries,
	seriesIndex: number,
	categoryLabels: ReadonlyArray<string>,
	categoryLevels: ReadonlyArray<ReadonlyArray<string>> | undefined,
	colorStride: number,
): HierarchyNode[] {
	const pointCount = series.values.length;
	const levels = normalizeLevels(categoryLevels ?? [], pointCount);
	const layout = series.parentLabelLayout ?? 'banner';
	const roots: HierarchyNode[] = [];
	for (let pointIndex = 0; pointIndex < pointCount; pointIndex++) {
		let children = roots;
		for (let levelIndex = levels.length - 1; levelIndex >= 1; levelIndex--) {
			const label = levels[levelIndex]?.[pointIndex] || `Group ${levelIndex}`;
			children = findOrAddParent(children, label, pointIndex, layout).children;
		}
		children.push({
			label: levels[0]?.[pointIndex] || categoryLabels[pointIndex] || String(pointIndex + 1),
			weight: Math.abs(series.values[pointIndex] ?? 0),
			seriesIndex,
			pointIndex,
			colorIndex: seriesIndex * colorStride + pointIndex,
			parentLabelLayout: layout,
		});
	}
	return roots;
}

/** Sets every parent's weight to the sum of its descendants'. */
export function aggregate(node: HierarchyNode): number {
	if (node.children) node.weight = node.children.reduce((sum, child) => sum + aggregate(child), 0);
	return node.weight;
}

/** Gives a top-level branch's colour index to every one of its descendants. */
function propagateColor(node: HierarchyNode, colorIndex: number): void {
	node.colorIndex = colorIndex;
	for (const child of node.children ?? []) propagateColor(child, colorIndex);
}

/**
 * The weighted tree of a chart: one root per top-level branch (one per series when there are
 * several), weights aggregated, and one colour index per top-level branch painted on every leaf
 * below it (COM-verified in PowerPoint: every leaf under "Branch 1" takes its legend swatch).
 */
export function buildHierarchyRoots(
	series: readonly HierarchySeries[],
	categoryLabels: ReadonlyArray<string>,
	categoryLevels: ReadonlyArray<ReadonlyArray<string>> | undefined,
): HierarchyNode[] {
	const colorStride = Math.max(categoryLabels.length, ...series.map((s) => s.values.length), 1);
	let roots = series.flatMap((s, index) =>
		buildSeriesHierarchy(s, index, categoryLabels, categoryLevels, colorStride),
	);
	if (series.length > 1)
		roots = series.map((s, index) => ({
			label: s.name,
			weight: 0,
			children: buildSeriesHierarchy(s, index, categoryLabels, categoryLevels, colorStride),
			colorIndex: index * colorStride,
			parentLabelLayout: s.parentLabelLayout ?? 'banner',
		}));
	for (const root of roots) aggregate(root);
	roots.forEach((root, index) => propagateColor(root, index));
	return roots;
}

/**
 * The top-level branch labels a legend should show, in the order colours are assigned to roots:
 * one per series when there are several, else one per unique outermost category level, else (a
 * flat treemap) one per leaf category.
 */
export function hierarchyBranchLabels(
	series: readonly { name: string; values: readonly number[] }[],
	categoryLabels: ReadonlyArray<string>,
	categoryLevels: ReadonlyArray<ReadonlyArray<string>> | undefined,
): string[] {
	if (series.length > 1) return series.map((s) => s.name);
	if (categoryLevels && categoryLevels.length > 1) {
		const pointCount = series[0]?.values.length ?? categoryLabels.length;
		const topLevel = normalizeLevels(categoryLevels, pointCount).at(-1) ?? [];
		const seen: string[] = [];
		for (const label of topLevel) if (!seen.includes(label)) seen.push(label);
		return seen;
	}
	return [...categoryLabels];
}

/** A rectangle of the treemap layout, in the order it is painted. */
export type TreemapCell =
	| { kind: 'parent'; label: string; x: number; y: number }
	| {
			kind: 'leaf';
			x: number;
			y: number;
			w: number;
			h: number;
			label: string;
			colorIndex: number;
			seriesIndex: number;
			pointIndex: number;
			/** Whether the cell is big enough to carry its label, and the label's font size. */
			labelFits: boolean;
			labelSize: number;
	  };

/** Height of a parent's banner label. */
const BANNER_HEIGHT = 16;

/** Squarified layout of the roots into `box`: parent labels and leaf cells, depth first. */
export function layoutTreemap(nodes: HierarchyNode[], box: TreemapBox): TreemapCell[] {
	const cells: TreemapCell[] = [];
	for (const [node, allocation] of squarify(nodes, box)) {
		const cell = {
			x: allocation.x + 1,
			y: allocation.y + 1,
			w: Math.max(allocation.w - 2, 1),
			h: Math.max(allocation.h - 2, 1),
		};
		if (node.children) {
			const showParent = node.parentLabelLayout !== 'none' && cell.w > 28 && cell.h > 16;
			if (showParent)
				cells.push({ kind: 'parent', label: node.label, x: cell.x + 4, y: cell.y + 11 });
			const banner = showParent && node.parentLabelLayout === 'banner' ? BANNER_HEIGHT : 0;
			cells.push(
				...layoutTreemap(node.children, {
					...cell,
					y: cell.y + banner,
					h: Math.max(cell.h - banner, 1),
				}),
			);
			continue;
		}
		cells.push({
			kind: 'leaf',
			...cell,
			label: node.label,
			colorIndex: node.colorIndex,
			seriesIndex: node.seriesIndex ?? 0,
			pointIndex: node.pointIndex ?? 0,
			labelFits: cell.w > 30 && cell.h > 14,
			labelSize: Math.min(10, cell.h * 0.3),
		});
	}
	return cells;
}
