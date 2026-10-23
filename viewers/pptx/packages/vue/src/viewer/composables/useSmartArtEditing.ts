import {
	rebuildDrawingShapesIfCleared,
	resolvePalette,
	smartArtLayoutSwitchPatch,
} from 'ooxml-ui/pptx';
import type { BoundingBox } from 'ooxml-ui/pptx';
import type {
	PptxElement,
	PptxSmartArtData,
	SmartArtColorScheme,
	SmartArtLayoutType,
	SmartArtStyle,
} from 'pptx-viewer-core';
import { computed } from 'vue';
import type { ComputedRef } from 'vue';

/**
 * useSmartArtEditing: editing logic for the Vue SmartArt inspector panel,
 * extracted out of the SFC so the component stays thin. Every mutation emits a
 * SHALLOW `Partial<PptxElement>` patch (`{ smartArtData }`) through the supplied
 * `apply` callback, matching the inspector-panel contract (`emit('update', patch)`
 * -> `useEditorOperations.updateElement`), so undo/redo keeps working. The text
 * pane edits come from the shared `pptx-ui-smartart-text-pane` element and land
 * through `commit`, like the layout, colour and style edits here.
 */

// Re-export the static option lists + layout-label helper from their dedicated
// module so existing callers can keep importing them from this barrel.
export {
	SMARTART_COLOR_SCHEMES,
	SMARTART_STYLE_OPTIONS,
	smartArtLayoutLabel,
} from './smartart-editing-constants';

export interface UseSmartArtEditingInput {
	/** Current SmartArt data for the selected element. */
	smartArtData: ComputedRef<PptxSmartArtData>;
	/** Apply a shallow element patch (typically `ops.updateElement`-backed). */
	apply: (patch: Partial<PptxElement>) => void;
	/**
	 * Owning element id + pixel box, used to reflow `drawingShapes` back from the
	 * layout engine whenever an edit clears them (see `rebuildDrawingShapesIfCleared`).
	 * Omit only in tests that don't care about post-edit rendering.
	 */
	elementId?: string;
	box?: ComputedRef<BoundingBox | undefined>;
}

export interface SmartArtEditingApi {
	colorScheme: ComputedRef<SmartArtColorScheme>;
	style: ComputedRef<SmartArtStyle>;
	currentLayout: ComputedRef<SmartArtLayoutType>;
	/** Commit a whole new SmartArt data object (e.g. a text pane edit). */
	commit: (next: PptxSmartArtData) => void;
	setColorScheme: (scheme: SmartArtColorScheme) => void;
	setStyle: (style: SmartArtStyle) => void;
	switchLayout: (layout: SmartArtLayoutType) => void;
}

export function useSmartArtEditing(input: UseSmartArtEditingInput): SmartArtEditingApi {
	const { smartArtData, apply } = input;

	const colorScheme = computed<SmartArtColorScheme>(
		() => smartArtData.value.colorScheme ?? 'colorful1',
	);
	const style = computed<SmartArtStyle>(() => smartArtData.value.style ?? 'flat');
	const currentLayout = computed<SmartArtLayoutType>(
		() => smartArtData.value.resolvedLayoutType ?? 'list',
	);

	/**
	 * Emit a whole new SmartArt-data object as the element patch, reflowing
	 * `drawingShapes` back from the layout engine first if the edit cleared them
	 * (every structural/text/style op does) -- otherwise the renderer falls back
	 * to the generic SVG layout for every node, not just the edited one.
	 */
	function commit(next: PptxSmartArtData): void {
		const box = input.box?.value;
		const reflowed = box
			? rebuildDrawingShapesIfCleared(
					next,
					next.layout,
					resolvePalette(next),
					next.style ?? 'flat',
					input.elementId ?? 'smartart',
					box,
				)
			: next;
		apply({ smartArtData: reflowed } as Partial<PptxElement>);
	}

	/** Emit a partial merge into the current SmartArt data. */
	function patchData(patch: Partial<PptxSmartArtData>): void {
		commit({ ...smartArtData.value, ...patch });
	}

	function setColorScheme(scheme: SmartArtColorScheme): void {
		if (scheme === colorScheme.value) {
			return;
		}
		patchData({ colorScheme: scheme });
	}

	function setStyle(next: SmartArtStyle): void {
		if (next === style.value) {
			return;
		}
		patchData({ style: next });
	}

	function switchLayout(layout: SmartArtLayoutType): void {
		if (layout === currentLayout.value) {
			return;
		}
		// Cleared fields (the old layoutDefinition, drawingShapes) must be forwarded
		// too, or the old layout keeps driving the relayout.
		patchData(smartArtLayoutSwitchPatch(smartArtData.value, layout));
	}

	return {
		colorScheme,
		style,
		currentLayout,
		commit,
		setColorScheme,
		setStyle,
		switchLayout,
	};
}
