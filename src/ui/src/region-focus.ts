/**
 * The F6 region service every Office product window shares.
 *
 * In Microsoft Office, F6 and Shift+F6 move the keyboard round the regions of the window (the
 * ribbon, the document, the task panes, the page or sheet tabs, the status bar) without walking
 * every control on the way. The products describe their regions as selectors in the order F6
 * visits them, and this module does the moving: it skips a region that is not on screen or has
 * nothing to focus, remembers nothing, and never traps the keyboard.
 *
 * Each product keeps its own region list and its own key dispatch, because a product knows which
 * pane is which; the walk itself lives here so there is one answer for every product and binding.
 *
 * @module region-focus
 */
import { containsComposed, modalActiveElement, modalFocusableElements } from './dialog/focus';

/** One region of a product window: a pane F6 stops on. */
export interface WindowRegion {
	/** Matches the region's container inside the scope; the first match is the region. */
	selector: string;
	/**
	 * Matches the stop inside the region that takes focus, instead of the region's own first
	 * stop: a message composer hands the keyboard to its text box, not to its Attach button.
	 */
	entry?: string;
	/** Takes focus itself and says so; false falls back to the entry stop. */
	focus?: (region: HTMLElement) => boolean;
}

/** A region the window is not showing at all: hidden, inert or hidden from the reader. */
const isClosed = (element: HTMLElement): boolean =>
	Boolean(element.hidden) || element.inert || element.getAttribute('aria-hidden') === 'true';

/** Where the regions are looked up: a document, a shadow root or an element of the window. */
export type RegionScope = Document | DocumentFragment | Element;

const ownerDocument = (scope: RegionScope): Document =>
	scope.nodeType === 9 ? (scope as Document) : ((scope as Element).ownerDocument ?? document);

/**
 * Where focus lands in a region: the stop `entry` names, else its current item (the selected tab,
 * the open page), else its first stop. A region with no stop of its own takes focus itself.
 */
export function regionEntryStop(region: HTMLElement, entry?: string): HTMLElement {
	const stops = modalFocusableElements(region);
	const named = entry ? stops.find((stop) => stop.matches(entry)) : undefined;
	const current = stops.find(
		(stop) =>
			stop.getAttribute('aria-selected') === 'true' ||
			(stop.getAttribute('aria-current') ?? 'false') !== 'false',
	);
	const stop = named ?? current ?? stops[0];
	if (stop) return stop;
	region.tabIndex = -1;
	return region;
}

/**
 * Moves focus to the next (`1`) or previous (`-1`) region of the window, as F6 and Shift+F6 do.
 * Returns false when the window shows no region to move to. With focus outside every region the
 * walk starts at the first region, or the last one going backwards, so the key always lands.
 */
export function moveRegionFocus(
	scope: RegionScope,
	regions: readonly WindowRegion[],
	direction: 1 | -1,
): boolean {
	const found = regions.flatMap((region) => {
		const element = scope.querySelector<HTMLElement>(region.selector);
		return element ? [{ region, element }] : [];
	});
	// A closed task pane is not a region, whether or not the page has been laid out.
	const open = found.filter(({ element }) => !isClosed(element));
	const shown = open.filter(({ element }) => element.getClientRects().length > 0);
	// Only laid-out regions count; with no layout at all (a test DOM) every one of them does.
	const list = shown.length > 0 ? shown : open;
	if (list.length === 0) return false;
	const doc = ownerDocument(scope);
	const at = list.findIndex(({ element }) => containsComposed(element, modalActiveElement(doc)));
	const from = at >= 0 ? at : direction > 0 ? -1 : list.length;
	// A region with nothing to focus is passed over, so F6 never stops on an empty pane.
	for (let step = 1; step <= list.length; step += 1) {
		const next = list[(from + direction * step + list.length * 2) % list.length]!;
		if (next.region.focus?.(next.element)) return true;
		regionEntryStop(next.element, next.region.entry).focus({ preventScroll: true });
		if (containsComposed(next.element, modalActiveElement(doc))) return true;
	}
	return false;
}

/** The direction a key press moves the regions: F6 forwards, Shift+F6 back, else nothing. */
export function regionKeyDirection(event: KeyboardEvent): 1 | -1 | undefined {
	if (event.key !== 'F6' || event.ctrlKey || event.metaKey || event.altKey || event.isComposing)
		return undefined;
	return event.shiftKey ? -1 : 1;
}
