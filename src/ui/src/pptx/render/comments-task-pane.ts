import { modalActiveElement } from '../../dialog/focus';

const DRAFT = 'textarea[data-pptx-comment-draft], [data-pptx-comment-draft] textarea';
const NEW_COMMENT =
	'[data-ribbon-control="insert.comments.comment"], [data-ribbon-control="review.comments.newComment"]';

/** Adds the shared close control and keyboard behavior to a docked comments surface. */
export function attachCommentsTaskPane(
	pane: HTMLElement,
	{ label, closeLabel, close }: { label: string; closeLabel: string; close: () => void },
): () => void {
	const doc = pane.ownerDocument;
	const root = pane.closest('[data-pptx-editor-chrome]') ?? doc;
	const opener = modalActiveElement(doc);
	pane.dataset.pptxCommentsPanel = 'true';
	pane.setAttribute('aria-label', label);
	const button = doc.createElement('button');
	button.type = 'button';
	button.dataset.pptxCommentClose = '';
	button.setAttribute('aria-label', closeLabel);
	button.title = closeLabel;
	const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 16 16');
	svg.setAttribute('aria-hidden', 'true');
	const path = doc.createElementNS(svg.namespaceURI, 'path');
	path.setAttribute('d', 'm4 4 8 8m0-8-8 8');
	svg.append(path);
	button.append(svg);
	pane.prepend(button);

	const focusDraft = () => (pane.querySelector<HTMLElement>(DRAFT) ?? button).focus();
	const frame = doc.defaultView?.requestAnimationFrame(focusDraft);
	const onCommand = (event: Event) => {
		if (event.composedPath().some((node) => node instanceof Element && node.matches(NEW_COMMENT))) {
			queueMicrotask(focusDraft);
		}
	};
	const onEscape = (event: KeyboardEvent) => {
		if (event.key === 'Escape' && !event.defaultPrevented) {
			event.preventDefault();
			event.stopPropagation();
			close();
		}
	};
	button.addEventListener('click', close);
	pane.addEventListener('keydown', onEscape);
	root.addEventListener('click', onCommand, true);
	return () => {
		if (frame !== undefined) doc.defaultView?.cancelAnimationFrame(frame);
		const active = modalActiveElement(doc);
		if (active instanceof Node && (pane.contains(active) || active === doc.body)) {
			queueMicrotask(() => {
				if (opener instanceof HTMLElement && opener.isConnected) {
					opener.focus();
					if (modalActiveElement(doc) === opener) return;
				}
				root
					.querySelector('[data-pptx-chrome="tab-row-actions"]')
					?.shadowRoot?.querySelector<HTMLElement>('[part~="comments"]')
					?.focus();
			});
		}
		button.removeEventListener('click', close);
		button.remove();
		pane.removeEventListener('keydown', onEscape);
		root.removeEventListener('click', onCommand, true);
		delete pane.dataset.pptxCommentsPanel;
	};
}
