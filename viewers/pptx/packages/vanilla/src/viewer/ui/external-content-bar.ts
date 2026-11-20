import { dispatchExternalContentBarIntent, registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxUiExternalContentBarElement } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';

/**
 * PowerPoint's security bar for linked pictures and media: "SECURITY WARNING
 * References to external pictures and other media have been blocked.
 * [Enable Content]". A thin adapter over the shared
 * `pptx-ui-external-content-bar`, which owns the markup and the strings.
 *
 * The bar is in the DOM only while it shows: `anchor` marks its place (after
 * the read-only banner), so a viewer without the bar has no trace of it.
 */
export interface ExternalContentBar {
	/** Where the bar goes when shown; the chrome appends it in the banner row. */
	anchor: Comment;
	update(visible: boolean): void;
}

export function createExternalContentBar(
	doc: Document,
	t: Translator,
	onEnable: () => void,
	onClose: () => void,
): ExternalContentBar {
	const anchor = doc.createComment('external-content-bar');
	let el: PptxUiExternalContentBarElement | null = null;
	return {
		anchor,
		update(visible) {
			if (!visible) {
				el?.remove();
				el = null;
				return;
			}
			if (!el) {
				registerPptxWebControls();
				el = doc.createElement('pptx-ui-external-content-bar');
				el.addEventListener('external-content-request', (event) =>
					dispatchExternalContentBarIntent(event, { enable: onEnable, close: onClose }),
				);
				el.state = { translate: t };
				anchor.after(el);
			}
		},
	};
}
