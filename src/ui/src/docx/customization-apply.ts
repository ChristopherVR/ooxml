/**
 * Applies `<docx-editor>`'s `customization` property: the capabilities reach the ribbon (through
 * `hiddenActions`), the read-only state, the File tab, the title bar, the context menu and the
 * shortcuts, which all read the resolved object from `core.viewOptions.customization`.
 */
import type { EditorCore } from './editor-core';
import { reflectAttribute } from './editor-attributes';
import {
	resolveWordCustomization,
	type WordCustomization,
	type WordRegionContext,
} from './customization';
import { applyViewOptions } from './view-options';

/**
 * Recomputes the effective read-only state: the `readOnly` property, or `capabilities.edit`
 * switched off. The attribute reflects the property only, so removing the capability gives the
 * host its own value back.
 */
export function applyReadOnly(core: EditorCore): void {
	core.readOnly = core.requestedReadOnly || !core.viewOptions.customization.canEdit;
	core.view?.setProps({ editable: () => core.canEditBody() });
	if (core.readOnly) core.parts.closeHeaderFooter(core.shell.canvas, core.shell.paper);
	core.refreshControls();
	reflectAttribute(core.element, 'read-only', core.requestedReadOnly);
}

/** Replaces the customisation (an equal object identity is a no-op) and re-applies it. */
export function setWordCustomization(core: EditorCore, value: WordCustomization | null): void {
	const { viewOptions } = core;
	if (value === viewOptions.customizationInput) return;
	const wasEditable = viewOptions.customization.canEdit;
	viewOptions.customizationInput = value;
	viewOptions.customization = resolveWordCustomization(value);
	applyViewOptions(core);
	if (viewOptions.customization.canEdit !== wasEditable) applyReadOnly(core);
	core.shell.chrome?.applyCustomization();
}

/** The context of a host region of `<docx-editor>` (a File tab page). */
export function wordRegionContext(core: EditorCore, region: string): WordRegionContext {
	return {
		region,
		editor: core.element,
		fileName: (core.element as HTMLElement & { fileName: string }).fileName,
		locale: core.locale,
		canEdit: !core.readOnly,
		close: () => core.shell.chrome?.closeBackstage(),
	};
}
