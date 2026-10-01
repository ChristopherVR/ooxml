/** Share parsed stylesheets between control instances of one document. */
const sheetsByDocument = new WeakMap<Document, Map<string, CSSStyleSheet>>();

/** Attach scoped CSS to a shadow root: a constructed stylesheet when available, else a `<style>`. */
export function attachStyles(root: ShadowRoot, css: string): void {
	const doc = root.host.ownerDocument;
	const Sheet = doc.defaultView?.CSSStyleSheet;
	if (Sheet && 'adoptedStyleSheets' in root && 'replaceSync' in Sheet.prototype) {
		try {
			let sheets = sheetsByDocument.get(doc);
			if (!sheets) {
				sheets = new Map();
				sheetsByDocument.set(doc, sheets);
			}
			let sheet = sheets.get(css);
			if (!sheet) {
				sheet = new Sheet();
				sheet.replaceSync(css);
				sheets.set(css, sheet);
			}
			root.adoptedStyleSheets = [...root.adoptedStyleSheets, sheet];
			return;
		} catch {
			// Test DOMs and older browsers can lack constructable stylesheet support.
		}
	}
	const style = doc.createElement('style');
	style.textContent = css;
	root.append(style);
}

/**
 * Rules every control repeats: `hidden`, the coarse-pointer touch target and the forced-colors
 * focus ring. Prepended to each control's own CSS by `controlCss`.
 */
export const COMMON_CONTROL_CSS = `
:host([hidden]) { display: none !important; }
@media (pointer: coarse), (max-width: 767px) {
	:host { --office-target-size: 44px; }
}
@media (forced-colors: active) {
	:host(:focus-visible), :focus-visible { outline: 2px solid Highlight; outline-offset: 2px; }
}
`;

export const controlCss = (css: string): string => `${COMMON_CONTROL_CSS}\n${css}`;
