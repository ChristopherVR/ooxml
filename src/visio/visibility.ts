import type { VisioLayer, VisioPage, VisioShape } from './model.js';
import { VisioPackageError } from './package-common.js';

/** Fixed work/allocation budgets, including for caller-constructed scenes. */
export const VISIO_VISIBILITY_LIMITS = Object.freeze({
	maxShapes: 25_000,
	maxDepth: 64,
	maxLayers: 25_000,
	maxMemberships: 100_000,
	maxMembershipsPerShape: 1024,
	maxLayerIdLength: 256,
});

export interface VisioVisibilityOptions {
	/** Exact IDs from this page's layers; values replace saved Visible, not Print. */
	layerVisibilityOverrides?: ReadonlyMap<string, boolean>;
}
export interface VisioResolvedShapeVisibility {
	/** Live input reference; the resolver never mutates or freezes the source shape. */
	readonly shape: VisioShape;
	/** Whole-shape suppression before ancestor suppression. */
	readonly ownHidden: boolean;
	readonly inheritedHidden: boolean;
	readonly hidden: boolean;
	/** Suppression from this shape's own memberships, excluding ancestor memberships. */
	readonly layerHidden: boolean;
}

function invalid(message: string): never {
	throw new VisioPackageError('INVALID_VISIBILITY_INPUT', message);
}
function bounded(count: number, maximum: number, name: string): void {
	if (!Number.isSafeInteger(count) || count < 0 || count > maximum)
		throw new VisioPackageError('VISIBILITY_LIMIT', `Visibility ${name} limit exceeded.`);
}
function layerId(id: string): void {
	if (typeof id !== 'string' || !id.length) invalid('Layer IDs must be nonempty strings.');
	bounded(id.length, VISIO_VISIBILITY_LIMITS.maxLayerIdLength, 'layer ID length');
}

/**
 * Resolve saved display visibility, optionally replacing page-local layer Visible flags.
 * Returns frozen preorder records in a frozen array. The source references stay live.
 *
 * No overrides preserves `hidden` exactly (plus existing ancestor suppression). With
 * overrides, only complete, consistent metadata may reveal a saved layer-hidden shape.
 * Legacy scenes or unknown NoShow caches can be hidden further but never revealed.
 * Missing layer references are ignored, matching parsing. Unknown override IDs reject.
 *
 * This is not a print resolver: Print, NonPrinting, Geometry.NoShow, HideText and group
 * own-data DisplayMode remain independent. DisplayMode=0 never hides group members.
 * Resolve background pages separately with their own layers and overrides.
 *
 * Display membership uses MS-VSDX 2.2.3.2.2: any invisible layer suppresses its shape.
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1b69f6b8-31d4-4b28-bc37-70ee4956d8cb
 */
export function resolveVisioPageVisibility(
	page: VisioPage,
	options: VisioVisibilityOptions = {},
): readonly VisioResolvedShapeVisibility[] {
	const limits = VISIO_VISIBILITY_LIMITS;
	const sourceLayers = page.layers ?? [];
	if (!Array.isArray(sourceLayers) || !Array.isArray(page.shapes))
		invalid('Page layers and shapes must be arrays.');
	bounded(sourceLayers.length, limits.maxLayers, 'layer count');
	const layers = new Map<string, VisioLayer>();
	for (const layer of sourceLayers) {
		if (!layer || typeof layer.visible !== 'boolean') invalid('Layer Visible must be boolean.');
		layerId(layer.id);
		// Match parser indexing when duplicate saved rows are present.
		if (!layers.has(layer.id)) layers.set(layer.id, layer);
	}
	const overrides = new Map<string, boolean>();
	if (options.layerVisibilityOverrides !== undefined) {
		const source = options.layerVisibilityOverrides;
		if (!source || typeof source[Symbol.iterator] !== 'function')
			invalid('Layer visibility overrides must be an iterable map.');
		bounded(source.size, limits.maxLayers, 'override count');
		let entries = 0;
		for (const [id, visible] of source) {
			bounded(++entries, limits.maxLayers, 'override count');
			layerId(id);
			if (!layers.has(id)) invalid('Layer visibility overrides must reference this page.');
			if (typeof visible !== 'boolean') invalid('Layer visibility overrides must be boolean.');
			overrides.set(id, visible);
		}
	}
	const result: VisioResolvedShapeVisibility[] = [];
	const seen = new Set<VisioShape>();
	const stack = [{ shapes: page.shapes, index: 0, inheritedHidden: false }];
	let memberships = 0;
	while (stack.length) {
		const frame = stack[stack.length - 1]!;
		if (frame.index >= frame.shapes.length) {
			stack.pop();
			continue;
		}
		bounded(result.length + frame.shapes.length - frame.index, limits.maxShapes, 'shape count');
		const shape = frame.shapes[frame.index++]!;
		if (!shape || typeof shape.hidden !== 'boolean' || !Array.isArray(shape.children))
			invalid('Each shape needs boolean hidden and an array of children.');
		if (seen.has(shape)) invalid('Shape trees cannot contain cycles or repeated shape objects.');
		seen.add(shape);
		const ids = shape.layerIds ?? [];
		if (!Array.isArray(ids)) invalid('Shape layer memberships must be an array.');
		bounded(ids.length, limits.maxMembershipsPerShape, 'shape membership count');
		memberships += ids.length;
		bounded(memberships, limits.maxMemberships, 'membership count');
		let savedLayerHidden = false,
			layerHidden = false;
		for (const id of ids) {
			layerId(id);
			const layer = layers.get(id);
			savedLayerHidden ||= layer?.visible === false;
			layerHidden ||= (overrides.get(id) ?? layer?.visible) === false;
		}
		let ownHidden = shape.hidden;
		if (overrides.size) {
			const reasons = shape.visibility;
			const complete =
				typeof reasons?.guide === 'boolean' &&
				typeof reasons.noShow === 'boolean' &&
				typeof reasons.layerHidden === 'boolean' &&
				reasons.layerHidden === savedLayerHidden &&
				shape.hidden === (reasons.guide || reasons.noShow || reasons.layerHidden);
			const intrinsicHidden = reasons?.guide === true || reasons?.noShow === true;
			ownHidden = layerHidden || intrinsicHidden || (!complete && shape.hidden);
		}
		const hidden = frame.inheritedHidden || ownHidden;
		result.push(
			Object.freeze({
				shape,
				ownHidden,
				inheritedHidden: frame.inheritedHidden,
				hidden,
				layerHidden,
			}),
		);
		if (shape.children.length) {
			bounded(stack.length, limits.maxDepth, 'shape depth');
			stack.push({ shapes: shape.children, index: 0, inheritedHidden: hidden });
		}
	}
	return Object.freeze(result);
}
