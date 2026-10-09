import { isVisioLayerName, type VisioAssignLayersEdit } from '../edit-layer-commands';
import type { VisioLayer, VisioPage } from '../model';

export interface VisioLayerAssignRow {
	layer: VisioLayer;
	/** Membership across the selected shapes, for a tri-state checkbox. */
	state: 'all' | 'some' | 'none';
}
export type VisioLayerAssignState =
	| { ok: true; rows: VisioLayerAssignRow[] }
	| { ok: false; reason: string };

/**
 * Scene check of Assign to Layer for the selected top-level shapes: no groups, no shapes on a
 * locked layer. Core admission (formulas, duplicate cells, page sheet) stays authoritative.
 */
export function visioLayerAssignState(
	page: VisioPage,
	shapeIds: readonly string[],
): VisioLayerAssignState {
	if (!shapeIds.length) return { ok: false, reason: 'Select one or more shapes.' };
	if (shapeIds.length > 1000 || new Set(shapeIds).size !== shapeIds.length)
		return { ok: false, reason: 'Select at most 1000 distinct shapes.' };
	const shapes = shapeIds.map((id) => page.shapes.find((shape) => shape.id === id));
	if (shapes.some((shape) => !shape)) return { ok: false, reason: 'Select top-level shapes.' };
	if (shapes.some((shape) => shape!.kind === 'group' || shape!.children.length))
		return { ok: false, reason: 'Groups cannot be assigned to layers here.' };
	const layers = page.layers ?? [];
	const locked = new Set(layers.filter((layer) => layer.locked).map((layer) => layer.id));
	if (shapes.some((shape) => shape!.layerIds?.some((id) => locked.has(id))))
		return { ok: false, reason: 'A selected shape is on a locked layer.' };
	const seen = new Set<string>();
	const rows: VisioLayerAssignRow[] = [];
	for (const layer of layers) {
		if (seen.has(layer.id)) continue;
		seen.add(layer.id);
		const count = shapes.filter((shape) => shape!.layerIds?.includes(layer.id)).length;
		rows.push({ layer, state: count === 0 ? 'none' : count === shapes.length ? 'all' : 'some' });
	}
	return { ok: true, rows };
}

/** Parse the dialog's New layer field: comma-separated names, trimmed, empty entries dropped. */
export function visioNewLayerNames(page: VisioPage, input: string): string[] | string {
	const names = input
		.split(',')
		.map((name) => name.trim())
		.filter(Boolean);
	const taken = new Set((page.layers ?? []).map((layer) => layer.name.toLowerCase()));
	const seen = new Set<string>();
	for (const name of names) {
		if (!isVisioLayerName(name))
			return `${name.slice(0, 40)} is not a valid layer name: use 1 to 255 characters without semicolons.`;
		const key = name.toLowerCase();
		if (taken.has(key) || seen.has(key)) return `A layer named ${name} already exists.`;
		seen.add(key);
	}
	return names;
}

/** The single source edit of the dialog: existing layer IDs plus any new layer names. */
export function visioLayerAssignCommand(
	page: VisioPage,
	shapeIds: readonly string[],
	layerIds: readonly string[],
	newLayers: readonly string[] = [],
): VisioAssignLayersEdit | undefined {
	const state = visioLayerAssignState(page, shapeIds);
	if (!state.ok) return undefined;
	const available = new Map(state.rows.map((row) => [row.layer.id, row.layer]));
	if (layerIds.some((id) => !available.has(id) || available.get(id)!.locked)) return undefined;
	return {
		type: 'assign-layers',
		pageId: page.id,
		shapeIds: [...shapeIds],
		layerIds: [...new Set(layerIds)],
		...(newLayers.length ? { newLayers: [...newLayers] } : {}),
	};
}
