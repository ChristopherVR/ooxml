import { armEditorKeyboard, buildRibbonGallery } from 'ooxml-ui/pptx';
import type {
	RibbonGalleryPickEvent,
	RibbonGalleryPlacement,
	RibbonGalleryPreviewEvent,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import type { RibbonGalleryHub, RibbonGalleryView } from './gallery-hub';

export interface RibbonGalleryOptions {
	chevronOnly?: boolean;
	/** Runs a command a gallery lists under its tiles ("Format Background..."). */
	onCommand?: (command: 'formatBackground') => void;
}
export interface RibbonGalleryControl extends RibbonGalleryView {
	el: HTMLElement;
	trigger: HTMLButtonElement;
	popup: HTMLElement;
	isOpen(): boolean;
}
/** View/CSS and popup lifecycle are shared; the hub retains native mutation/history. */
export function createRibbonGallery(
	doc: Document,
	t: Translator,
	placement: RibbonGalleryPlacement,
	hub: RibbonGalleryHub,
	options: RibbonGalleryOptions = {},
): RibbonGalleryControl {
	const el = doc.createElement('pptx-ui-ribbon-gallery');
	el.setAttribute('mode', placement.mode);
	el.toggleAttribute('chevron-only', Boolean(options.chevronOnly));
	if (!options.chevronOnly) {
		el.setAttribute('data-ribbon-control', placement.control);
	}
	el.translateLabel = (key, params) => t(key, params ? { ...params } : undefined);
	el.addEventListener('gallery-pick', (event) => {
		if (!el.disabled) {
			const { itemId, action } = (event as RibbonGalleryPickEvent).detail;
			const command = hub.pick(placement.gallery, itemId, action);
			if (command?.kind === 'command') {
				options.onCommand?.(command.command);
			}
			armEditorKeyboard(el.closest<HTMLElement>('.pptxv'));
		}
	});
	el.addEventListener('gallery-preview', (event) => {
		if (!el.disabled) {
			hub.preview(placement.gallery, (event as RibbonGalleryPreviewEvent).detail.itemId);
		}
	});
	el.addEventListener('gallery-preview-end', () => hub.endPreview());
	const control: RibbonGalleryControl = {
		el,
		// Live reads: the element re-renders its trigger, so a node captured here goes stale.
		get trigger() {
			return el.trigger;
		},
		get popup() {
			return el.popup;
		},
		isOpen: () => el.open,
		close: () => el.close(),
		refresh(ctx, editable) {
			el.descriptor = buildRibbonGallery(placement.gallery, ctx);
			el.disabled = !editable;
		},
	};
	hub.register(control);
	return control;
}
