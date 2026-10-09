import { visioShapeFormattingState } from 'ooxml-core/visio/ui';
import type { VisioShape } from 'ooxml-core/visio';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { RibbonCommand } from './ribbon-parts';
import { effectsOptions, QUICK_STYLES_HINT, quickStyleGalleryState } from './ribbon-shape-styles';

/**
 * Selection state for Home > Shape Styles: the Quick Styles gallery and the Effects > Shadow
 * presets follow the same admission as Fill and Line. Unsupported effects keep their reasons.
 */
export function renderShapeStyleControls(
	root: ShadowRoot,
	shapes: readonly VisioShape[],
	reason: string,
	set: (element: RibbonCommand | null, reason: string) => void,
): void {
	const gallery = root.querySelector<OfficeUiGallery>(
		'office-ui-gallery[data-menu="quick-styles"]',
	);
	if (gallery) {
		const disabled = !!reason;
		if (gallery.state?.disabled !== disabled) gallery.state = quickStyleGalleryState(disabled);
		gallery.toggleAttribute('disabled', disabled);
		gallery.title = disabled ? `Quick Styles: ${reason}` : QUICK_STYLES_HINT;
	}
	set(root.querySelector('[data-menu="effects"]'), reason);
	const shadow = effectsOptions().find((item) => item.id === 'shadow');
	set(root.querySelector('[data-menu="shadow"]'), reason);
	const current = reason ? undefined : visioShapeFormattingState(shapes).shadowPreset;
	for (const item of shadow?.items ?? []) {
		const element = root.querySelector<RibbonCommand>(`[command="${item.id}"]`);
		set(element, reason);
		if (element && item.action?.type === 'shape-format')
			element.setAttribute('checked', String(current === item.action.patch.shadow));
	}
}
