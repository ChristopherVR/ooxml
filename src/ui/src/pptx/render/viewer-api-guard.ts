/**
 * The edit methods of the imperative viewer handle, gated by the viewer's effective edit right.
 *
 * A host `canEdit={false}`, a read-only cloud file, Protected View, a read-only recommendation and
 * `setMode('preview')` all end up as one boolean in each binding. The UI honours it by disabling
 * or hiding commands, but the handle is a second way in: a host (or an AI bridge) calling
 * `addSlide` or `updateElement` on a read-only deck must be refused the same way `addElement` and
 * `updateElements` already refuse. Every binding wraps its handle with {@link gateViewerEdits}, so
 * none can forget one of the methods.
 *
 * @module render/viewer-api-guard
 */
import type { PowerPointViewerAPI, ViewerMode } from '../types';

/** The handle methods that change the document and do not check the edit right themselves. */
export const GATED_VIEWER_EDIT_METHODS = [
	'addSlide',
	'deleteSlides',
	'duplicateSlides',
	'moveSlide',
	'toggleHideSlides',
	'updateElement',
	'deleteElements',
	'duplicateElement',
] as const;

export type GatedViewerEditMethod = (typeof GATED_VIEWER_EDIT_METHODS)[number];

type GatedViewerEditApi = Pick<PowerPointViewerAPI, GatedViewerEditMethod>;

/** A deck with no slides: there is nothing to protect yet. */
export function isEmptyDeck(api: Pick<PowerPointViewerAPI, 'getSlideCount'>): boolean {
	return api.getSlideCount() === 0;
}

/**
 * `api` with its edit methods turned into no-ops (`duplicateElement` returns `undefined`) while
 * `canEdit()` is false. `canEdit` is read on every call, so a later editable document restores
 * them without rebuilding the handle. Other members are returned untouched.
 *
 * One exception: `addSlide` still works on a deck with no slides at all, so a host can create the
 * first slide of an empty viewer it opened read-only (host-side deck creation); every other edit,
 * and `addSlide` on a deck that has slides, is refused.
 */
export function gateViewerEdits<
	T extends GatedViewerEditApi & Pick<PowerPointViewerAPI, 'getSlideCount'>,
>(api: T, canEdit: () => boolean): T {
	const gated: Record<string, unknown> = { ...api };
	for (const name of GATED_VIEWER_EDIT_METHODS) {
		const method = api[name] as (...args: unknown[]) => unknown;
		gated[name] = (...args: unknown[]) =>
			canEdit() || (name === 'addSlide' && isEmptyDeck(api)) ? method.apply(api, args) : undefined;
	}
	return gated as T;
}

/**
 * The mode `getMode()` reports: a viewer that may not edit is in `preview`, whatever mode was last
 * requested (`present` and `master` stay as they are).
 */
export function resolveReportedViewerMode(mode: ViewerMode, canEdit: boolean): ViewerMode {
	return mode === 'edit' && !canEdit ? 'preview' : mode;
}
