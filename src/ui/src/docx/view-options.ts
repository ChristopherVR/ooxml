import type { EditorCore } from './editor-core';
import { findLocalizedControl } from './localization';
import type { RibbonActionId } from 'ooxml-core/docx/ui';
import { applyRibbonVisibility } from './ribbon-visibility';
import { mergeHiddenActions } from 'ooxml-core/embedding';
import {
	EMPTY_WORD_CUSTOMIZATION,
	type ResolvedWordCustomization,
	type WordCustomization,
} from './customization';

/** Host-controlled UI customisation; each value maps to one element property and attribute. */
export class ViewOptions {
	showThumbnails = false;
	showToolbar = true;
	hiddenActions: readonly RibbonActionId[] = [];
	/** The host's `customization` object, as given. */
	customizationInput: WordCustomization | null = null;
	/** `customizationInput` resolved for the render sites. */
	customization: ResolvedWordCustomization = EMPTY_WORD_CUSTOMIZATION;
}

/** Pushes the current options into the built shell; a no-op until the element connects. */
export function applyViewOptions(core: EditorCore): void {
	const { toolbar, navigator } = core.shell;
	const { showThumbnails, showToolbar, hiddenActions, customization } = core.viewOptions;
	navigator?.setOpen(showThumbnails);
	if (!toolbar) return;
	applyRibbonVisibility(
		toolbar,
		showToolbar,
		mergeHiddenActions(hiddenActions, customization.hiddenActions),
	);
	findLocalizedControl(toolbar, 'Page thumbnails')?.setAttribute(
		'aria-pressed',
		String(showThumbnails),
	);
}
