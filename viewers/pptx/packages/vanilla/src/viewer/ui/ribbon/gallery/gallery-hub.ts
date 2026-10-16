import type {
	RibbonGalleryApplyResult,
	RibbonGalleryContext,
	RibbonGalleryId,
} from 'ooxml-ui/pptx';
import { applyRibbonGalleryItem } from 'ooxml-ui/pptx';

/** A mounted gallery control the hub keeps in step with the selection. */
export interface RibbonGalleryView {
	refresh(ctx: RibbonGalleryContext, editable: boolean): void;
	close(): void;
}

/**
 * The one place the ribbon's galleries read their context from and send their
 * picks through. The ribbon pushes the selection-derived context in (push-style,
 * like every other Home tab control), and each gallery control registers
 * itself once at construction; that keeps the context out of every group's own
 * `update` signature.
 */
export interface RibbonGalleryHub {
	context(): RibbonGalleryContext;
	register(view: RibbonGalleryView): void;
	/** Rebuild every registered gallery when the selection, theme or editability changed. */
	sync(ctx: RibbonGalleryContext | undefined, editable: boolean): void;
	/**
	 * Resolve a tile pick (with the context menu `action` chosen on the tile, if any) through the
	 * shared registry and dispatch the result. Returns it when it is a command the caller runs.
	 */
	pick(gallery: RibbonGalleryId, itemId: string, action?: string): RibbonGalleryApplyResult | null;
	/** Show what a previewable tile would do while hovered; `endPreview` puts the slides back. */
	preview(gallery: RibbonGalleryId, itemId: string): void;
	endPreview(): void;
	closeAll(): void;
}

const EMPTY_CONTEXT: RibbonGalleryContext = { element: null };

/** Hover preview hooks of the Design galleries. */
export interface RibbonGalleryPreviewHooks {
	preview(result: RibbonGalleryApplyResult | null): void;
	end(): void;
}

export function createRibbonGalleryHub(
	dispatch: (result: RibbonGalleryApplyResult) => void,
	previewHooks?: RibbonGalleryPreviewHooks,
): RibbonGalleryHub {
	const views: RibbonGalleryView[] = [];
	let ctx: RibbonGalleryContext = EMPTY_CONTEXT;
	let editable = false;
	let key: readonly unknown[] = [];

	const refreshAll = (): void => {
		for (const view of views) {
			view.refresh(ctx, editable);
		}
	};

	return {
		context: () => ctx,
		register(view) {
			views.push(view);
			view.refresh(ctx, editable);
		},
		sync(next, nextEditable) {
			const resolved = next ?? EMPTY_CONTEXT;
			// The chrome sync runs on every store change; building a gallery
			// serialises dozens of preview SVGs, so only rebuild when something a
			// descriptor reads actually changed.
			const nextKey = [
				resolved.element,
				resolved.theme?.colorScheme,
				resolved.theme?.fontScheme,
				resolved.theme?.name,
				resolved.theme?.formatScheme?.name,
				resolved.activeSlide?.backgroundStyleRef,
				resolved.activeSlide?.clrMapOverride,
				resolved.themeVariants,
				resolved.resolveStyleMatrix !== undefined,
				// By value: the context builds a fresh array on every sync (Table Layout reads it).
				resolved.selectedTableCells?.map((cell) => `${cell.row}:${cell.column}`).join(','),
				nextEditable,
			];
			ctx = resolved;
			editable = nextEditable;
			if (nextKey.length === key.length && nextKey.every((value, index) => value === key[index])) {
				return;
			}
			key = nextKey;
			refreshAll();
		},
		pick(gallery, itemId, action) {
			const result = applyRibbonGalleryItem(gallery, itemId, ctx, action);
			if (result?.kind === 'command') {
				previewHooks?.end();
				return result;
			}
			if (result) {
				dispatch(result);
			}
			return null;
		},
		preview(gallery, itemId) {
			previewHooks?.preview(applyRibbonGalleryItem(gallery, itemId, ctx));
		},
		endPreview() {
			previewHooks?.end();
		},
		closeAll() {
			for (const view of views) {
				view.close();
			}
		},
	};
}
