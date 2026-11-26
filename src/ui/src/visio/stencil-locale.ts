import { translateRibbonText, type VisioLocale } from './ribbon-locale';
import { ribbonLocale } from './ribbon-localize';

let strings: Readonly<Record<VisioLocale, Readonly<Record<string, string>>>> | undefined;

export async function loadStencilStrings(): Promise<void> {
	strings ??= (await import('./locales/stencils')).VISIO_STENCIL_STRINGS;
}

export function translateStencilText(pane: HTMLElement, text: string): string {
	const locale = ribbonLocale(pane);
	const dictionary = strings?.[locale];
	return dictionary && Object.hasOwn(dictionary, text)
		? (dictionary[text] ?? text)
		: translateRibbonText(locale, text);
}
