import { defineElement } from '../../registry';
import { defineFormatSection } from '../../controls';
import { base, bridged } from './office-aliases';

/**
 * `pptx-ui-format-section`: the shared `office-ui-format-section` (one
 * collapsible Format pane section) with the pptx token bridge, so it follows
 * the viewer's theme. Same properties and `office-section-toggle` event.
 */
export function definePptxFormatSection(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-format-section')) {
		return;
	}
	defineElement(
		registry,
		'pptx-ui-format-section',
		class extends bridged(
			base(registry, defineFormatSection, 'office-ui-format-section'),
			'pptx-ui-format-section',
		) {},
	);
}
