import { normalizeVisioLocale, translateRibbonText, type VisioLocale } from './ribbon-locale';

/** The attributes that carry user-visible wording on the ribbon's elements. */
const ATTRIBUTES = [
	'label',
	'title',
	'aria-label',
	'launcher-label',
	'more-label',
	'none-label',
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
 * Translates the ribbon in place: tab names, group and command labels, tooltips, accessible names
 * and the shared ribbon's own File and collapse labels. Add-in tabs keep the host's wording.
 * Calling it again with another locale translates from the English originals.
 */
export function localizeRibbon(ribbon: HTMLElement, value: string | null | undefined): VisioLocale {
	const locale = normalizeVisioLocale(value);
	ribbon.dataset.visioLocale = locale;
	const elements = [ribbon, ...ribbon.querySelectorAll<HTMLElement>('*')];
	for (const element of elements) {
		if (element.closest('[data-add-in]')) continue;
		for (const name of ATTRIBUTES) {
			const text = element.getAttribute(name);
			if (text)
				translateSlot(element, name, text, locale, (next) => element.setAttribute(name, next));
		}
		if (element.childElementCount === 0 && element.textContent?.trim()) {
			const text = element.textContent;
			translateSlot(element, '#text', text, locale, (next) => (element.textContent = next));
		}
	}
	ribbon.setAttribute('file-label', translateRibbonText(locale, 'File'));
	ribbon.setAttribute('collapse-label', translateRibbonText(locale, 'Collapse the ribbon'));
	return locale;
}
