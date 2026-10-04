import { tok } from './tokens.js';

/** Scope of every light-DOM gallery rule; the element marks itself on connection. */
const H = '[data-office-gallery]';

/**
 * Light-DOM styles: the gallery renders its tiles beside the host's own children so products can
 * query them, so the rules live in the host's root (document or shadow root), once per root.
 */
export const GALLERY_LIGHT_CSS = `
${H} { display: inline-flex; position: relative; flex: none; color: ${tok('--office-foreground')}; }
${H} .gallery-view { display: contents; }
${H} button { box-sizing: border-box; font: inherit; color: inherit; cursor: pointer;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius')}; background: ${tok('--office-background')}; }
${H} button:hover:not(:disabled) { background: ${tok('--office-selected')}; }
${H} button:disabled { opacity: .4; cursor: not-allowed; }
${H} button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
${H} .trigger { display: inline-flex; align-items: center; gap: ${tok('--office-space-1-5')}; min-height: ${tok('--office-gallery-trigger-height')};
	padding: 0 ${tok('--office-space-1-5')}; font-size: ${tok('--office-font-size-sm')}; white-space: nowrap; }
${H}:not([mode=inline]) .trigger { border-color: transparent; background: transparent; }
${H} .trigger svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; fill: none; stroke: currentColor;
	stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; }
${H}[mode=inline] .trigger, ${H}[chevron-only] .trigger { align-self: stretch; padding: 0 ${tok('--office-space-1')}; }
${H} .strip { display: flex; gap: ${tok('--office-space-0')}; padding: ${tok('--office-space-0')}; align-items: center; }
${H} .tile { display: inline-flex; align-items: center; justify-content: center; flex: none; padding: ${tok('--office-space-0')}; }
${H} .tile[aria-pressed=true] { border: ${tok('--office-border-width-thick')} solid ${tok('--office-accent')};
	background: color-mix(in srgb, ${tok('--office-accent')} 15%, ${tok('--office-background')}); }
${H} .tile svg { max-width: 100%; max-height: 100%; }
${H}[mode=inline] .strip { gap: 0; border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-inline-end: 0;
	border-radius: ${tok('--office-radius-sm')} 0 0 ${tok('--office-radius-sm')}; background: ${tok('--office-gallery-strip-background')}; }
${H}[mode=inline] .trigger { border-radius: 0 ${tok('--office-radius-sm')} ${tok('--office-radius-sm')} 0; background: ${tok('--office-gallery-strip-background')}; }
${H} .trigger.command { border-color: transparent; background: transparent; justify-content: flex-start; }
${H} .trigger.command:hover:not(:disabled) { border-color: ${tok('--office-border')}; background: ${tok('--office-selected')}; }
${H} .trigger.command svg { color: ${tok('--office-command-icon-color')}; }
${H} .trigger.command-large { flex-direction: column; align-items: center; justify-content: flex-start; gap: ${tok('--office-space-0')};
	min-width: ${tok('--office-gallery-command-min-width')}; max-width: ${tok('--office-gallery-command-max-width')}; height: ${tok('--office-gallery-command-height')};
	padding: ${tok('--office-space-1')}; font-size: ${tok('--office-font-size-2xs')}; line-height: ${tok('--office-line-height-tight')}; white-space: normal; text-align: center; }
${H} .trigger.command-large svg { width: ${tok('--office-gallery-command-icon')}; height: ${tok('--office-gallery-command-icon')}; }
${H} .popup { box-sizing: border-box; position: fixed; z-index: ${tok('--office-z-popover')}; max-width: calc(100vw - ${tok('--office-space-4')});
	max-height: ${tok('--office-gallery-popup-max-height')}; overflow: auto; padding: ${tok('--office-space-2')};
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-md')};
	background: ${tok('--office-popover')}; color: ${tok('--office-popover-foreground')}; box-shadow: ${tok('--office-shadow-lg')}; }
${H} .popup[hidden] { display: none; }
${H} .grid { display: grid; gap: ${tok('--office-space-1')}; width: max-content; max-width: 100%; }
${H} .heading { margin: ${tok('--office-space-1')} 0 ${tok('--office-space-1-5')}; font-size: ${tok('--office-font-size-xs')}; font-weight: ${tok('--office-font-weight-bold')}; }
${H} .section + .section { margin-top: ${tok('--office-space-2-5')}; }
@media (pointer: coarse), (max-width: 1023px) {
	${H} .trigger, ${H} .tile { min-height: ${tok('--office-target-size-touch')}; min-width: ${tok('--office-target-size-touch')}; }
}
@media (forced-colors: active) {
	${H} button, ${H} .popup { color: ButtonText; background: Canvas; border-color: ButtonText; }
	${H} button:disabled { color: GrayText; opacity: 1; }
	${H} button:focus-visible, ${H} .tile[aria-pressed=true] { outline: ${tok('--office-focus-width')} solid Highlight; border-color: Highlight; }
}
`;

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
