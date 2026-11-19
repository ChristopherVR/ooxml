import { defineElement } from '../../registry';
import { defineOptionsDialog } from '../../controls';
import { base, bridged } from './office-aliases';

/**
 * `pptx-ui-options-dialog`: the shared `office-ui-options-dialog` under the pptx token bridge, so
 * File > Options follows the viewer theme. Every binding mounts it; see
 * `render/options/options-dialog-binding` for how it is driven.
 */
export function definePptxOptionsDialog(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-options-dialog')) {
		return;
	}
	defineElement(
		registry,
		'pptx-ui-options-dialog',
		class extends bridged(
			base(registry, defineOptionsDialog, 'office-ui-options-dialog'),
			'pptx-ui-options-dialog',
		) {},
	);
}
