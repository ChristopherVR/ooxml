import type { SignatureStatusView } from 'ooxml-ui/pptx';

import { createEl } from '../render';
import { createIcon } from './icons';

/**
 * The chrome's signature badge, a row under the Ribbon where React draws its
 * toolbar badge. It shows the shared `signatureStatusView` descriptor, so a
 * deck reads "Signed" only when its signatures verify, and renders no button
 * while the deck carries no signature parts. A click opens File > Info >
 * Digital Signatures.
 */
export interface SignatureBadge {
	el: HTMLElement;
	update(view: SignatureStatusView): void;
}

export function createSignatureBadge(doc: Document, onOpen: () => void): SignatureBadge {
	const el = createEl(doc, 'div', 'pptxv-signature-badge-row');
	el.hidden = true;
	return {
		el,
		update(view) {
			el.hidden = !view.present;
			if (!view.present) {
				el.replaceChildren();
				return;
			}
			const button = createEl(doc, 'button', 'pptxv-signature-badge');
			button.type = 'button';
			button.dataset.pptxSignatureState = view.state;
			button.dataset.tone = view.tone;
			button.title = view.description;
			button.setAttribute('aria-label', view.description);
			const label = createEl(doc, 'span');
			label.textContent = view.label;
			button.append(createIcon(doc, view.icon), label);
			button.addEventListener('click', onOpen);
			el.replaceChildren(button);
		},
	};
}
