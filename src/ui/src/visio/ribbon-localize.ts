import { normalizeVisioLocale, translateRibbonText, type VisioLocale } from './ribbon-locale';

/** The attributes that carry user-visible wording on the ribbon's elements. */
const ATTRIBUTES = [
	'label',
	'title',
	'aria-label',
	'launcher-label',
	'more-label',
	'none-label',
	'empty-label',
	'placeholder',
	'data-label',
] as const;

interface Applied {
	/** The English wording the translation came from. */
	english: string;
	/** What was written, to tell a later English update from our own output. */
	written: string;
}

const applied = new WeakMap<Element, Map<string, Applied>>();

/** Sets `name` on `element` to the translation of its English wording and remembers the original. */
function translateSlot(
	element: Element,
	name: string,
	current: string,
	locale: VisioLocale,
	write: (text: string) => void,
): void {
	const slots = applied.get(element) ?? new Map<string, Applied>();
	applied.set(element, slots);
	const before = slots.get(name);
	// A value the product wrote since our last pass is new English wording.
	const english = before && before.written === current ? before.english : current;
	const written = translateRibbonText(locale, english);
	slots.set(name, { english, written });
	if (written !== current) write(written);
}

/** The locale of a ribbon last localized with `localizeRibbon`. */
export function ribbonLocale(ribbon: Element): VisioLocale {
	return normalizeVisioLocale((ribbon as HTMLElement).dataset.visioLocale);
}

/**
 * Translates the wording of `elements` in place: their attributes in `ATTRIBUTES` and the text of
 * their leaf elements. Add-in content and `data-user-text` (a document's own names) are kept.
 * Calling it again with another locale translates from the English originals.
 */
export function localizeWording(elements: Iterable<Element>, locale: VisioLocale): void {
	for (const element of elements) {
		if (element.closest('[data-add-in], [data-user-text]')) continue;
		for (const name of ATTRIBUTES) {
			const text = element.getAttribute(name);
			if (text)
				translateSlot(element, name, text, locale, (next) => element.setAttribute(name, next));
		}
		// A gallery's tiles are drawn by Lit into the gallery itself: rewriting their text would
		// remove the markers Lit renders into, and its next update would throw.
		if (
			element.childElementCount === 0 &&
			element.textContent?.trim() &&
			!element.closest('office-ui-gallery')
		) {
			const text = element.textContent;
			translateSlot(element, '#text', text, locale, (next) => (element.textContent = next));
		}
	}
}

/**
 * Translates the ribbon in place: tab names, group and command labels, tooltips, accessible names
 * and the shared ribbon's own File and collapse labels. Add-in tabs keep the host's wording.
 * Calling it again with another locale translates from the English originals.
 */
export function localizeRibbon(ribbon: HTMLElement, value: string | null | undefined): VisioLocale {
	const locale = normalizeVisioLocale(value);
	ribbon.dataset.visioLocale = locale;
	localizeWording([ribbon, ...ribbon.querySelectorAll<HTMLElement>('*')], locale);
	ribbon.setAttribute('file-label', translateRibbonText(locale, 'File'));
	ribbon.setAttribute('collapse-label', translateRibbonText(locale, 'Collapse the ribbon'));
	return locale;
}
