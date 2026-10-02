import {
	getVisioPageLayers,
	resolveVisioPageVisibility,
	VISIO_VISIBILITY_LIMITS,
	type VisioDocument,
	type VisioShape,
} from 'ooxml-core/visio';

/** UI limits. Visibility semantics and shape traversal remain in ooxml-core. */
export const VIEWER_LAYER_LIMITS = Object.freeze({
	overrides: VISIO_VISIBILITY_LIMITS.maxLayers,
	controls: 200,
});
export interface LayerVisibilityOverride {
	/** The source page owning the layer, including background pages. */
	readonly pageId: string;
	readonly layerId: string;
	readonly visible: boolean;
}
export const EMPTY_LAYER_OVERRIDES: readonly LayerVisibilityOverride[] = Object.freeze([]);

/** Validate before allocating indexes; never expose mutable maps as viewer state. */
export function layerOverrideMaps(
	model: VisioDocument | null,
	overrides: readonly LayerVisibilityOverride[],
): Map<string, Map<string, boolean>> {
	if (!Array.isArray(overrides) || overrides.length > VIEWER_LAYER_LIMITS.overrides)
		throw new Error('The viewer exceeds the layer override limit.');
	const pages = new Map(model?.pages.map((page) => [page.id, page]));
	const known = new Map<string, Set<string>>();
	const result = new Map<string, Map<string, boolean>>();
	for (const entry of overrides) {
		if (
			!entry ||
			typeof entry.pageId !== 'string' ||
			typeof entry.layerId !== 'string' ||
			typeof entry.visible !== 'boolean'
		)
			throw new Error('Layer visibility overrides require page and layer IDs and a boolean.');
		const page = pages.get(entry.pageId);
		if (!page) throw new Error('The layer override page does not belong to this document.');
		let ids = known.get(page.id);
		if (!ids) known.set(page.id, (ids = new Set(page.layers?.map((layer) => layer.id))));
		if (!ids.has(entry.layerId)) throw new Error('The layer does not belong to this page.');
		let values = result.get(page.id);
		if (!values) result.set(page.id, (values = new Map()));
		if (values.has(entry.layerId)) throw new Error('Duplicate layer visibility override.');
		values.set(entry.layerId, entry.visible);
	}
	return result;
}

export function documentVisibility(
	model: VisioDocument | null,
	overrides: readonly LayerVisibilityOverride[] = EMPTY_LAYER_OVERRIDES,
): WeakMap<VisioShape, boolean> {
	const maps = layerOverrideMaps(model, overrides);
	const visible = new WeakMap<VisioShape, boolean>();
	for (const page of model?.pages ?? []) {
		const layerVisibilityOverrides = maps.get(page.id);
		for (const entry of resolveVisioPageVisibility(
			page,
			layerVisibilityOverrides ? { layerVisibilityOverrides } : {},
		))
			visible.set(entry.shape, !entry.hidden);
	}
	return visible;
}

export function visibleSelection(
	model: VisioDocument | null,
	pageIndex: number,
	selection: { readonly id: string; readonly pageId?: string } | null,
	visible: WeakMap<VisioShape, boolean>,
): boolean {
	const page = model?.pages[pageIndex];
	if (!model || !page || !selection) return false;
	const source = getVisioPageLayers(model, page.id).find(
		(candidate) => candidate.id === (selection.pageId ?? page.id),
	);
	const stack = [...(source?.shapes ?? [])];
	while (stack.length) {
		const shape = stack.pop()!;
		if (shape.id === selection.id) return hasVisibleShapeContent(shape, visible);
		stack.push(...shape.children);
	}
	return false;
}

/** Rendering eligibility shared by selection and SVG containers; no format interpretation. */
export function hasVisibleShapeContent(
	shape: VisioShape,
	visible: WeakMap<VisioShape, boolean> | undefined,
	cache = new WeakMap<VisioShape, boolean>(),
): boolean {
	const cached = cache.get(shape);
	if (cached !== undefined) return cached;
	const own =
		!(shape.kind === 'group' && shape.groupDisplayMode === 0) &&
		(shape.geometry.length > 0 ||
			!!shape.image ||
			!!shape.text.plainText ||
			shape.kind === 'foreign');
	const renderable =
		(visible?.get(shape) ?? !shape.hidden) &&
		(own || shape.children.some((child) => hasVisibleShapeContent(child, visible, cache)));
	cache.set(shape, renderable);
	return renderable;
}
