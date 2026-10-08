import { getVisioPageLayers } from '../parser';
import { VISIO_VISIBILITY_LIMITS } from '../visibility';
import type { VisioDocument, VisioShape } from '../model';
import type { VisioShapeSelection } from './contract';

export const EMPTY_SELECTION: readonly VisioShapeSelection[] = Object.freeze([]);
export const selectionKey = (shape: Pick<VisioShapeSelection, 'id' | 'pageId'>, pageId: string) =>
	JSON.stringify([shape.pageId ?? pageId, shape.id]);

/** Omitted source page means the displayed page; background identities remain distinct. */
export function visioSelectionIsOnPage(
	selection: Pick<VisioShapeSelection, 'pageId'> | null | undefined,
	pageId: string,
): boolean {
	return !!selection && (selection.pageId === undefined || selection.pageId === pageId);
}
export function sameSelection(
	left: readonly VisioShapeSelection[],
	right: readonly VisioShapeSelection[],
): boolean {
	return (
		left.length === right.length &&
		left.every((shape, index) => {
			const other = right[index]!;
			return shape.id === other.id && shape.pageId === other.pageId && shape.name === other.name;
		})
	);
}

/** Snapshot displayed targets once, and keep selected groups from also selecting descendants. */
export function snapshotSelection(
	model: VisioDocument | null,
	pageIndex: number,
	input: readonly VisioShapeSelection[],
	visible: WeakMap<VisioShape, boolean>,
): readonly VisioShapeSelection[] {
	if (!Array.isArray(input)) throw new Error('Selection requires an array.');
	const length = input.length;
	if (length > VISIO_VISIBILITY_LIMITS.maxShapes)
		throw new Error('Selection exceeds the supported shape count.');
	const page = model?.pages[pageIndex];
	const candidates = new Map<string, VisioShapeSelection>();
	// Read bounded scalar records once; host iterators and later getter mutations cannot leak in.
	for (let index = 0; index < length; index++) {
		const item = input[index];
		const id = item?.id,
			name = item?.name,
			pageId = item?.pageId;
		if (
			typeof id !== 'string' ||
			typeof name !== 'string' ||
			id.length > 1024 ||
			name.length > 4096 ||
			(pageId !== undefined && (typeof pageId !== 'string' || pageId.length > 256))
		)
			throw new Error('Selection requires shape IDs, names and optional page IDs.');
		const copied = Object.freeze({ id, name, ...(pageId === undefined ? {} : { pageId }) });
		const key = selectionKey(copied, page?.id ?? '');
		if (!candidates.has(key)) candidates.set(key, copied);
	}
	if (!model || !page || !length) return EMPTY_SELECTION;
	const accepted = new Map<string, readonly string[]>();
	const content = new WeakMap<VisioShape, boolean>();
	for (const source of getVisioPageLayers(model, page.id)) {
		const pending = source.shapes.map((shape) => ({ shape, ancestors: [] as readonly string[] }));
		while (pending.length) {
			const { shape, ancestors } = pending.pop()!;
			const key = selectionKey({ id: shape.id, pageId: source.id }, page.id);
			if (candidates.has(key) && hasVisibleShapeContent(shape, visible, content))
				accepted.set(key, ancestors);
			for (const child of shape.children)
				pending.push({ shape: child, ancestors: [...ancestors, key] });
		}
	}
	const result: VisioShapeSelection[] = [];
	for (const [key, item] of candidates) {
		const ancestors = accepted.get(key);
		if (!ancestors || ancestors.some((ancestor) => accepted.has(ancestor))) continue;
		result.push(item);
	}
	return result.length ? Object.freeze(result) : EMPTY_SELECTION;
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
			!!shape.foreignVector ||
			!!shape.text.plainText ||
			shape.kind === 'foreign');
	const renderable =
		(visible?.get(shape) ?? !shape.hidden) &&
		(own || shape.children.some((child) => hasVisibleShapeContent(child, visible, cache)));
	cache.set(shape, renderable);
	return renderable;
}
