import { translate, type EditorLocale, type LocalizationKey } from '../localization';

import { enReferences } from '../locales/en-references';
import { deReferences } from '../locales/de-references';
import { frReferences } from '../locales/fr-references';
import { esReferences } from '../locales/es-references';
import { zhCNReferences } from '../locales/zh-CN-references';
import type { ReferencesLabel } from '../locales/en-references';

export type ReferencesText = ReferencesLabel | LocalizationKey;

const strings = {
	en: enReferences,
	de: deReferences,
	fr: frReferences,
	es: esReferences,
	'zh-CN': zhCNReferences,
};

const isReferencesLabel = (key: ReferencesText): key is ReferencesLabel =>
	Object.hasOwn(enReferences, key);

export function referencesText(locale: EditorLocale, key: ReferencesText): string {
	return isReferencesLabel(key) ? strings[locale][key] : translate(locale, key);
}
