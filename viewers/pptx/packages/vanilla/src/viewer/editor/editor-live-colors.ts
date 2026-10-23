import {
	createLiveColorPatchInput,
	fillColorOf,
	shapeFillChange,
	shapeOutlineChange,
	strokeColorOf,
	textColorOf,
	textStylePatch,
} from 'ooxml-ui/pptx';
import type { LiveColorInput } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import type { ApplyToSelected } from './editor-apply-to-selected';
import { patchShapeStyle } from './editor-format-mutations';
import type { EditorOps } from './editor-operations';
import { recordRecentColor } from './editor-recent-colors';

/** Which of the inspector's native colour pickers an event comes from. */
export type LiveColorTarget = 'fill' | 'line' | 'text';

/** The native colour input events a live picker reports. */
export type LiveColorEventType = 'input' | 'change' | 'blur';

/**
 * The inspector's native colour pickers, live while the user drags (shared
 * `createLiveColorPatchInput`): every shade is merged into the selected
 * element without history, and the settled pick (`change`, or `blur` for a
 * drag no change ended) puts the previews back and commits once through
 * `applyToSelected`, so a drag is one undo step and one recent colour. A pick
 * that settles where it started is no undo step (its `change` still counts as
 * a recent colour).
 */
export interface LiveColorActions {
	/** One `input`, `change` or `blur` of the fill, line or text colour picker. */
	liveColor(target: LiveColorTarget, type: LiveColorEventType, hex: string): void;
	/** Show an effects colour drag's shape-style frame without an undo step. */
	previewShapeStyle(patch: Partial<ShapeStyle>): void;
}

export function createLiveColorActions(deps: {
	store: Store<ViewerState>;
	ops: EditorOps;
	applyToSelected: ApplyToSelected;
}): LiveColorActions {
	const { store, ops, applyToSelected } = deps;
	const editable = (): PptxElement | undefined =>
		store.get().editable ? ops.selectedElement() : undefined;

	const liveInput = (
		colorOf: (el: PptxElement) => string,
		patch: (el: PptxElement, hex: string) => Partial<PptxElement>,
	): LiveColorInput =>
		createLiveColorPatchInput<PptxElement, Partial<PptxElement>>({
			read: () => {
				const el = editable();
				return el && { key: el.id, color: colorOf(el), state: el };
			},
			patch,
			preview: (next) => {
				const el = editable();
				if (el) {
					ops.previewElement(el.id, next);
				}
			},
			commit: (next) => applyToSelected(() => next),
			remember: (hex) => recordRecentColor(store, hex),
		});

	// A native pick is a custom colour: it clears any theme ref, like every
	// commit path that does not come from a theme swatch.
	const inputs: Record<LiveColorTarget, LiveColorInput> = {
		fill: liveInput(fillColorOf, (el, hex) => patchShapeStyle(el, shapeFillChange(hex, undefined))),
		line: liveInput(strokeColorOf, (el, hex) =>
			patchShapeStyle(
				el,
				shapeOutlineChange(hex, undefined, hasShapeProperties(el) ? (el.shapeStyle ?? {}) : {}),
			),
		),
		text: liveInput(textColorOf, (el, hex) =>
			'textStyle' in el ? textStylePatch(el, { color: hex, colorRef: undefined }) : {},
		),
	};

	return {
		liveColor(target, type, hex) {
			inputs[target][type](hex);
		},
		previewShapeStyle(patch) {
			const el = editable();
			if (el) {
				ops.previewElement(el.id, patchShapeStyle(el, patch));
			}
		},
	};
}
