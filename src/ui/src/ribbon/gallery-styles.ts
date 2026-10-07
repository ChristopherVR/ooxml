import { withTokens } from '../base';
import css from './gallery.css?raw';

/**
 * The gallery renders its tiles in the light DOM beside the host's own children so products can
 * query them, so its rules live in the host's root (a document or a shadow root), once per root.
 */
const GALLERY_LIGHT_CSS = withTokens(css);

const injected = new WeakSet<Node>();

/** Add the light-DOM rules to the root holding `host` (a document or a shadow root), once. */
export function attachGalleryStyles(host: HTMLElement): void {
	const root = host.getRootNode() as Document | ShadowRoot;
	if (injected.has(root)) return;
	const doc = host.ownerDocument;
	const Sheet = doc.defaultView?.CSSStyleSheet;
	if (Sheet && 'adoptedStyleSheets' in root && 'replaceSync' in Sheet.prototype) {
		try {
			const sheet = new Sheet();
			sheet.replaceSync(GALLERY_LIGHT_CSS);
			root.adoptedStyleSheets = [...root.adoptedStyleSheets, sheet];
			injected.add(root);
			return;
		} catch {
			// Fall back to a style element below.
		}
	}
	const style = doc.createElement('style');
	style.dataset.officeGalleryStyles = '';
	style.textContent = GALLERY_LIGHT_CSS;
	(root.nodeType === 9 ? doc.head : root).append(style);
	injected.add(root);
}
