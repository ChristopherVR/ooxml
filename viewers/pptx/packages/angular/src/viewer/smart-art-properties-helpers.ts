/**
 * smart-art-properties-helpers.ts: pure logic for the Angular SmartArt inspector.
 *
 * The presentational `SmartArtPropertiesComponent` is intentionally thin: every
 * mutation it performs is delegated to one of the pure functions below, which in
 * turn wrap the framework-agnostic SmartArt editing operations re-exported from
 * `pptx-viewer-core` and `ooxml-ui/pptx`. The text pane's edits come from the
 * shared `pptx-ui-smartart-text-pane` element and need no helper here.
 *
 * Mirrors the React inspector:
 *   packages/react/src/viewer/components/inspector/SmartArtPropertiesPanel.tsx
 *   packages/react/src/viewer/components/inspector/SmartArtLayoutSwitcher.tsx
 *
 * @module angular-viewer/smart-art-properties-helpers
 */

import { setSmartArtInspectorColor, switchSmartArtLayoutData } from 'ooxml-ui/pptx';
import type { PptxSmartArtData, SmartArtLayoutType, SmartArtStyle } from 'pptx-viewer-core';

// ── Option constants (mirror the React panel) ───────────────────────────────

/** Selectable style intensities, in display order. */
export const SMART_ART_STYLE_OPTIONS: readonly SmartArtStyle[] = [
	'flat',
	'moderate',
	'intense',
] as const;

/** Default style used when the data has none set. */
export const DEFAULT_STYLE: SmartArtStyle = 'flat';

/** Default resolved layout used when the data has none set. */
export const DEFAULT_LAYOUT: SmartArtLayoutType = 'list';

// ── Read helpers ─────────────────────────────────────────────────────────────

/** Current style intensity, falling back to the default. */
export function currentStyle(data: PptxSmartArtData): SmartArtStyle {
	return data.style ?? DEFAULT_STYLE;
}

/** Current resolved layout, falling back to the default. */
export function currentLayout(data: PptxSmartArtData): SmartArtLayoutType {
	return data.resolvedLayoutType ?? DEFAULT_LAYOUT;
}

// ── Mutation helpers (all return new immutable PptxSmartArtData) ──────────────

export { setSmartArtInspectorColor as setColorScheme };

/** Apply a new style intensity; clears stale drawing shapes for reflow. */
export function setStyle(data: PptxSmartArtData, style: SmartArtStyle): PptxSmartArtData {
	return { ...data, style, drawingShapes: undefined };
}

/**
 * Switch the diagram to a new layout, preserving node data and connections.
 * Returns the original data unchanged when the target equals the current layout.
 */
export function setLayout(data: PptxSmartArtData, layout: SmartArtLayoutType): PptxSmartArtData {
	if (currentLayout(data) === layout) {
		return data;
	}
	return switchSmartArtLayoutData(data, layout);
}
