import {
	VISIO_STENCILS,
	visioOutlineShape,
	type VisioBasicOutline,
	type VisioBasicShape,
	type VisioOutlineShape,
} from 'ooxml-core/visio';

export interface Master {
	id: string;
	name: string;
	/** Preview outline in a 24x24 box (static application geometry). */
	path: string;
	/** Default size in inches when dropped. */
	size: { width: number; height: number };
}
export interface Stencil {
	id: string;
	name: string;
	masters: readonly Master[];
}
/** How the core creates a master: a native ellipse, or a core outline. */
export type MasterCreation = { kind: 'ellipse' } | { kind: 'rectangle'; shape: VisioOutlineShape };

const ELLIPSES = new Set(['ellipse', 'circle']);
const unit = { width: 1, height: 1 };
const wide = { width: 1, height: 0.75 };
/** Visio's Basic Shapes stencil, in Visio's order. */
export const BASIC_SHAPES: readonly Master[] = [
	{ id: 'rectangle', name: 'Rectangle', path: 'M3 6h18v12H3Z', size: wide },
	{ id: 'square', name: 'Square', path: 'M5 4h14v14H5Z', size: unit },
	{ id: 'ellipse', name: 'Ellipse', path: 'M2 12a10 6.5 0 1 0 20 0 10 6.5 0 1 0-20 0', size: wide },
	{ id: 'circle', name: 'Circle', path: 'M4 12a8 8 0 1 0 16 0 8 8 0 1 0-16 0', size: unit },
	{ id: 'triangle', name: 'Triangle', path: 'M12 4 21 20H3Z', size: unit },
	{ id: 'right-triangle', name: 'Right triangle', path: 'M4 4v16h16Z', size: unit },
	{ id: 'pentagon', name: 'Pentagon', path: 'M12 3l9 6.5-3.4 10.5H6.4L3 9.5Z', size: unit },
	{ id: 'hexagon', name: 'Hexagon', path: 'M7 4h10l5 8-5 8H7l-5-8Z', size: wide },
	{ id: 'octagon', name: 'Octagon', path: 'M8.5 3h7L21 8.5v7L15.5 21h-7L3 15.5v-7Z', size: unit },
	{
		id: 'star',
		name: '5-point star',
		path: 'm12 2.5 2.8 6.4 6.9.6-5.2 4.6 1.6 6.8L12 17.3l-6.1 3.6 1.6-6.8-5.2-4.6 6.9-.6Z',
		size: unit,
	},
	{ id: 'diamond', name: 'Diamond', path: 'M12 2.5 21.5 12 12 21.5 2.5 12Z', size: unit },
	{
		id: 'rounded-rectangle',
		name: 'Rounded rectangle',
		path: 'M6 6h12a3 3 0 0 1 3 3v6a3 3 0 0 1-3 3H6a3 3 0 0 1-3-3V9a3 3 0 0 1 3-3Z',
		size: wide,
	},
	{ id: 'cross', name: 'Cross', path: 'M9 3h6v6h6v6h-6v6H9v-6H3V9h6Z', size: unit },
	{ id: 'parallelogram', name: 'Parallelogram', path: 'M7 6h15l-5 12H2Z', size: wide },
	{ id: 'trapezoid', name: 'Trapezoid', path: 'M7 6h10l5 12H2Z', size: wide },
	{
		id: 'can',
		name: 'Can',
		path: 'M5 6a7 2.5 0 1 0 14 0 7 2.5 0 1 0-14 0v12a7 2.5 0 0 0 14 0V6',
		size: { width: 0.75, height: 1 },
	},
	{ id: 'cube', name: 'Cube', path: 'M4 8h12v12H4ZM4 8l4-4h12l-4 4M20 4v12l-4 4', size: unit },
	{ id: 'chevron', name: 'Chevron', path: 'M3 5h13l5 7-5 7H3l5-7Z', size: wide },
];
export const BASIC_STENCIL_ID = 'basic';

const n = (value: number) => String(Math.round(value * 100) / 100);
/** A preview path for a core outline (or a native ellipse) drawn at its drop proportions. */
export function outlinePreviewPath(
	outline: VisioBasicOutline | 'ellipse',
	size: { width: number; height: number },
): string {
	const scale = 20 / Math.max(size.width, size.height);
	const w = size.width * scale,
		h = size.height * scale;
	const left = 12 - w / 2,
		top = 12 - h / 2;
	if (outline === 'ellipse')
		return `M${n(left)} 12a${n(w / 2)} ${n(h / 2)} 0 1 0 ${n(w)} 0a${n(w / 2)} ${n(h / 2)} 0 1 0 ${n(-w)} 0`;
	if (outline.rounding && outline.paths.length === 1 && outline.paths[0]!.length === 4) {
		const r = outline.rounding * Math.min(w, h);
		return `M${n(left + r)} ${n(top)}h${n(w - 2 * r)}a${n(r)} ${n(r)} 0 0 1 ${n(r)} ${n(r)}v${n(h - 2 * r)}a${n(r)} ${n(r)} 0 0 1 ${n(-r)} ${n(r)}h${n(2 * r - w)}a${n(r)} ${n(r)} 0 0 1 ${n(-r)} ${n(-r)}v${n(2 * r - h)}a${n(r)} ${n(r)} 0 0 1 ${n(r)} ${n(-r)}Z`;
	}
	const point = ([x, y]: readonly [number, number]) => `${n(left + x * w)} ${n(top + (1 - y) * h)}`;
	return [
		...outline.paths.map((points) => `M${points.map(point).join('L')}Z`),
		...(outline.lines ?? []).map((points) => `M${points.map(point).join('L')}`),
	].join('');
}

/** Basic Shapes, then the built-in stencils More Shapes can open, with generated previews. */
export const STENCILS: readonly Stencil[] = [
	{ id: BASIC_STENCIL_ID, name: 'Basic Shapes', masters: BASIC_SHAPES },
	...VISIO_STENCILS.map((stencil) => ({
		id: stencil.id,
		name: stencil.name,
		masters: stencil.masters.map((master) => ({
			id: master.id,
			name: master.name,
			size: master.size,
			path: outlinePreviewPath(
				master.shape === 'ellipse' ? 'ellipse' : visioOutlineShape(master.shape),
				master.size,
			),
		})),
	})),
];
const ELLIPSE_MASTERS = new Set(
	VISIO_STENCILS.flatMap((stencil) => stencil.masters)
		.filter((master) => master.shape === 'ellipse')
		.map((master) => master.id),
);
const OUTLINES = new Map(
	VISIO_STENCILS.flatMap((stencil) => stencil.masters).map((master) => [master.id, master.shape]),
);

/** The master and its stencil for a master id, or nothing for an unknown id. */
export function findMaster(id: string): { stencil: Stencil; master: Master } | undefined {
	for (const stencil of STENCILS) {
		const master = stencil.masters.find((candidate) => candidate.id === id);
		if (master) return { stencil, master };
	}
	return undefined;
}

/** Default size and creation of a stencil master, or nothing for an unknown id. */
export function masterCreation(
	id: string,
): { size: { width: number; height: number }; create: MasterCreation } | undefined {
	const found = findMaster(id);
	if (!found) return undefined;
	const create: MasterCreation =
		ELLIPSES.has(id) || ELLIPSE_MASTERS.has(id)
			? { kind: 'ellipse' }
			: {
					kind: 'rectangle',
					shape:
						found.stencil.id === BASIC_STENCIL_ID
							? (id as VisioBasicShape)
							: (OUTLINES.get(id) as VisioOutlineShape),
				};
	return { size: found.master.size, create };
}
