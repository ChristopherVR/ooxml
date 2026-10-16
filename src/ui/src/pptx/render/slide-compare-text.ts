/**
 * The sentence a Compare panel shows for one {@link ElementChange}, spelled in the viewer's
 * language. `slide-compare` keeps the English sentence in `description`; every binding shows
 * `describeElementChange(change, t)` instead, so the wording is decided once.
 *
 * @module render/slide-compare-text
 */
import type { ElementChange } from './slide-compare';

type Translate = (key: string, params?: Record<string, string | number>) => string;

const capitalise = (word: string): string => word.charAt(0).toUpperCase() + word.slice(1);

function lookup(translate: Translate, key: string, fallback: string): string {
	const value = translate(key);
	return value && value !== key ? value : fallback;
}

/** "Shape (id)" or `Text: "preview"` with the element type in the viewer's language. */
function elementLabel(
	element: NonNullable<ElementChange['descriptionElement']>,
	translate: Translate,
): string | undefined {
	const type = lookup(
		translate,
		`pptx.compare.elementType.${element.type}`,
		capitalise(element.type),
	);
	const key =
		element.text === undefined ? 'pptx.compare.elementLabel' : 'pptx.compare.elementLabelText';
	const label = translate(key, { type, id: element.id, text: element.text ?? '' });
	return label && label !== key ? label : undefined;
}

/** The change's description in the active language; the English sentence when no key resolves. */
export function describeElementChange(change: ElementChange, translate: Translate): string {
	if (!change.descriptionKey) {
		return change.description;
	}
	const params: Record<string, string | number> = { ...change.descriptionParams };
	if (change.descriptionElement) {
		const label = elementLabel(change.descriptionElement, translate);
		if (label === undefined) {
			return change.description;
		}
		params.label = label;
	}
	for (const name of ['from', 'to']) {
		if (change.descriptionKey === 'pptx.compare.change.background' && params[name] === '') {
			params[name] = lookup(translate, 'pptx.compare.defaultValue', 'default');
		}
	}
	const text = translate(change.descriptionKey, params);
	return text && text !== change.descriptionKey ? text : change.description;
}
