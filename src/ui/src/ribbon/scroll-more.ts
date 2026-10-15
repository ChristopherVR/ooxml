import css from './scroll-more.css?raw';

const STYLE_ID = 'office-ui-ribbon-touch-style';
const CLASS = 'ribbon-scroll-more';

/**
 * Adds the shared touch-target and "more" chevron rules once to the tree that holds `node`: its
 * shadow root when an editor renders inside one, otherwise the document head.
 */
export function ensureRibbonTouchStyle(node: Node): void {
	const root = node.getRootNode();
	const doc = node.ownerDocument ?? (node as Document);
	const host: ParentNode | null = root instanceof ShadowRoot ? root : doc.head;
	if (!host || host.querySelector(`#${STYLE_ID}`)) return;
	const style = doc.createElement('style');
	style.id = STYLE_ID;
	style.textContent = css;
	host.append(style);
}

/**
 * The accessible name of the chevron: a fixed string, or a function of the panel it sits in. The
 * function is called on every sync, so a locale change applies without rebuilding the options.
 */
export type ScrollMoreLabel = string | ((panel: HTMLElement) => string);

/**
 * Pins a "more" chevron to the end of `panel` while its content still overflows (even after
 * every group is folded). It sticks inside the visible area; a press scrolls the panel on by one
 * page and wraps back to the start at the end.
 */
export function syncScrollMore(panel: HTMLElement, label: ScrollMoreLabel = 'More'): void {
	panel.querySelector(`:scope > .${CLASS}`)?.remove();
	if (panel.hidden || panel.scrollWidth <= panel.clientWidth + 1) return;
	ensureRibbonTouchStyle(panel);
	const doc = panel.ownerDocument;
	const text = typeof label === 'function' ? label(panel) : label;
	const button = doc.createElement('button');
	button.type = 'button';
	button.className = CLASS;
	button.setAttribute('aria-label', text);
	button.title = text;
	button.textContent = '›';
	button.addEventListener('mousedown', (event) => event.preventDefault());
	button.addEventListener('click', () => {
		const end = panel.scrollWidth - panel.clientWidth - 2;
		const left = panel.scrollLeft >= end ? 0 : panel.scrollLeft + panel.clientWidth * 0.8;
		panel.scrollTo({ left, behavior: 'smooth' });
	});
	panel.append(button);
}
