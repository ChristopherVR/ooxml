import { ribbonLocale } from './ribbon-localize';
import { translateRibbonText } from './ribbon-locale';

/**
 * A cue that the ribbon's tab strip scrolls. The shared `office-ui-ribbon` already snaps its tabs
 * and keeps the active one in view; when more tabs lie past either edge this marks the ribbon
 * (`data-tabs-start` / `data-tabs-end`, which the stylesheet turns into a fade on that edge) and
 * keeps a chevron button in the ribbon's `end` slot that scrolls the strip by a page, wrapping
 * back to the first tab from the end. Returns a disposer.
 */
export function wireRibbonTabCue(ribbon: HTMLElement): () => void {
	const doc = ribbon.ownerDocument;
	const view = doc.defaultView;
	const button = doc.createElement('button');
	button.type = 'button';
	button.className = 'tab-scroll-more';
	button.slot = 'end';
	button.hidden = true;
	button.addEventListener('mousedown', (event) => event.preventDefault());
	ribbon.append(button);
	let list: HTMLElement | null = null;
	let frame = 0;
	const sync = () => {
		frame = 0;
		const strip = list;
		const overflow = !!strip && strip.scrollWidth - strip.clientWidth > 1;
		const atStart = !strip || strip.scrollLeft <= 1;
		const atEnd = !strip || strip.scrollLeft >= strip.scrollWidth - strip.clientWidth - 1;
		ribbon.toggleAttribute('data-tabs-start', overflow && !atStart);
		ribbon.toggleAttribute('data-tabs-end', overflow && !atEnd);
		button.hidden = !overflow;
		const text = translateRibbonText(ribbonLocale(ribbon), 'More tabs');
		button.setAttribute('aria-label', text);
		button.title = text;
		button.textContent = overflow && atEnd ? '‹' : '›';
	};
	const schedule = () => {
		if (!frame && view) frame = view.requestAnimationFrame(sync);
	};
	const resize = view && 'ResizeObserver' in view ? new view.ResizeObserver(schedule) : undefined;
	/** The tab strip is drawn inside the ribbon's shadow root, after its first render. */
	const attach = () => {
		const found = ribbon.shadowRoot?.querySelector<HTMLElement>('[role="tablist"]') ?? null;
		if (found === list) return;
		list?.removeEventListener('scroll', schedule);
		resize?.disconnect();
		list = found;
		list?.addEventListener('scroll', schedule, { passive: true });
		if (list) resize?.observe(list);
		resize?.observe(ribbon);
		schedule();
	};
	button.addEventListener('click', () => {
		const strip = list;
		if (!strip) return;
		const end = strip.scrollWidth - strip.clientWidth - 2;
		const left = strip.scrollLeft >= end ? 0 : strip.scrollLeft + strip.clientWidth * 0.8;
		strip.scrollTo({ left, behavior: 'smooth' });
	});
	const ready = (ribbon as HTMLElement & { updateComplete?: Promise<unknown> }).updateComplete;
	void ready?.then(attach);
	// Tabs added or relabelled (a locale change, an add-in) rebuild the strip's content.
	const mutation =
		view && 'MutationObserver' in view
			? new view.MutationObserver(() => (attach(), schedule()))
			: undefined;
	mutation?.observe(ribbon, {
		attributes: true,
		subtree: true,
		attributeFilter: ['data-label', 'data-tab-hidden'],
	});
	attach();
	return () => {
		if (frame) view?.cancelAnimationFrame(frame);
		list?.removeEventListener('scroll', schedule);
		resize?.disconnect();
		mutation?.disconnect();
		button.remove();
	};
}
