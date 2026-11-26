import { translateRibbonText, type VisioLocale } from './ribbon-locale';
import { ribbonLocale } from './ribbon-localize';
import { VISIO_SHAPES_STRINGS } from './locales/shapes';

let strings: Readonly<Record<VisioLocale, Readonly<Record<string, string>>>> | undefined;

export async function loadStencilStrings(): Promise<void> {
	strings ??= (await import('./locales/stencils')).VISIO_STENCIL_STRINGS;
}

export function translateStencilText(pane: HTMLElement, text: string): string {
	return stencilText(ribbonLocale(pane), text);
}

export function stencilText(locale: VisioLocale, text: string): string {
	const shapes: Readonly<Record<string, string>> = VISIO_SHAPES_STRINGS[locale];
	if (Object.hasOwn(shapes, text)) return shapes[text] ?? text;
	const dictionary = strings?.[locale];
	return dictionary && Object.hasOwn(dictionary, text)
		? (dictionary[text] ?? text)
		: translateRibbonText(locale, text);
}
