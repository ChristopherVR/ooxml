/**
 * Applies `<docx-editor>`'s `customization` property: the capabilities reach the ribbon (through
 * `hiddenActions`), the read-only state (a lock beside the host's switch and a viewer role), the
 * File tab, the title bar, the context menu and the shortcuts, which all read the resolved object
 * from `core.viewOptions.customization`.
 */
import type { EditorCore } from './editor-core';
import {
	resolveWordCustomization,
	type WordCustomization,
	type WordRegionContext,
} from './customization';
import { applyViewOptions } from './view-options';

/** Replaces the customisation (an equal object identity is a no-op) and re-applies it. */
export function setWordCustomization(core: EditorCore, value: WordCustomization | null): void {
	const { viewOptions } = core;
	if (value === viewOptions.customizationInput) return;
	// `capabilities.edit` is one of the read-only locks (`EditorCore.readOnly`); reporting through
	// the core keeps `readonly-change` on the combined value and never touches the host's switch.
	core.reportReadOnly(() => {
		viewOptions.customizationInput = value;
		viewOptions.customization = resolveWordCustomization(value);
	});
	applyViewOptions(core);
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
