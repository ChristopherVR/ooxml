/**
 * Lazy access to PowerPoint's built-in SmartArt layout definitions.
 *
 * The library (`pptx-viewer-core/smartart-layouts`, about 350 KB gzipped) is a separate chunk
 * loaded the first time a SmartArt is selected, so a viewer that never shows one never fetches it.
 * Until it has loaded, callers get `undefined` and keep the family-renderer behaviour; the layout
 * gallery and the switch patch both re-check on every call, so they pick it up as soon as it lands.
 *
 * @module render/smartart-builtin-layouts
 */
import type { PptxSmartArtData } from 'pptx-viewer-core';

type BuiltinLayouts = typeof import('pptx-viewer-core/smartart-layouts');

let loaded: BuiltinLayouts | undefined;
let loading: Promise<BuiltinLayouts | undefined> | undefined;

/** Start loading the library (once); resolves with it, or `undefined` when it cannot load. */
export function preloadSmartArtBuiltinLayouts(): Promise<BuiltinLayouts | undefined> {
	loading ??= import('pptx-viewer-core/smartart-layouts').then(
		(module) => (loaded = module),
		() => undefined,
	);
	return loading;
}

/** The library when it has finished loading, else `undefined` (and a load is started). */
export function smartArtBuiltinLayouts(): BuiltinLayouts | undefined {
	if (!loaded) {
		void preloadSmartArtBuiltinLayouts();
	}
	return loaded;
}

/** The data fields a built-in layout swap changes, as a patch (see `smartArtLayoutSwitchPatch`). */
export function builtinLayoutPatch(
	data: PptxSmartArtData,
	layoutId: string,
): Partial<PptxSmartArtData> | undefined {
	const library = smartArtBuiltinLayouts();
	if (!library?.findBuiltinSmartArtLayout(layoutId)) {
		return undefined;
	}
	const next = library.applyBuiltinSmartArtLayout(data, layoutId);
	return {
		layoutType: next.layoutType,
		resolvedLayoutType: next.resolvedLayoutType,
		layout: next.layout,
		layoutDefinition: next.layoutDefinition,
		builtinLayoutId: next.builtinLayoutId,
		presLayoutVars: next.presLayoutVars,
		layoutDirty: next.layoutDirty,
		drawingDirty: next.drawingDirty,
		drawingShapes: next.drawingShapes,
	};
}
