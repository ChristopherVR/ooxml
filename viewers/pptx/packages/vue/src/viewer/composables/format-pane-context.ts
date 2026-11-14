import { createFormatPaneController } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';
import { inject } from 'vue';
import type { InjectionKey } from 'vue';

/**
 * The viewer's Format pane controller (open sections, launch requests), one per
 * `PowerPointViewer` instance. A component mounted on its own (tests, a host
 * composing the inspector) gets a private one.
 */
export const FormatPaneKey: InjectionKey<FormatPaneController> = Symbol('pptx-format-pane');

/** The provided controller, or a private one when none is provided. */
export function injectFormatPane(): FormatPaneController {
	return inject(FormatPaneKey, null) ?? createFormatPaneController();
}
