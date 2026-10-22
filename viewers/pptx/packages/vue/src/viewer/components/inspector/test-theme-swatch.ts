import { DOMWrapper } from '@vue/test-utils';

/** A mounted component or one of its elements. */
interface Finder {
	find(selector: string): DOMWrapper<Element>;
}

/** The shadow root of the shared Theme Colors grid (`pptx-ui-color-grid`) under `wrapper`. */
export function themeColorGrid(wrapper: Finder): ShadowRoot | null {
	const grid = wrapper.find('pptx-ui-color-grid');
	return grid.exists() ? grid.element.shadowRoot : null;
}

/** A Theme Colors swatch of that grid by its label, wrapped so a test can `trigger` it. */
export function themeSwatch(wrapper: Finder, title: string): DOMWrapper<HTMLButtonElement> {
	const swatch = themeColorGrid(wrapper)?.querySelector<HTMLButtonElement>(
		`button.swatch[title="${title}"]`,
	);
	if (!swatch) {
		throw new Error(`No Theme Colors swatch titled "${title}"`);
	}
	return new DOMWrapper(swatch);
}
