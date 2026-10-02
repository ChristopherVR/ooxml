// Translation tables of the command strings: English key -> [fr, de, es, zh-CN]. Each dialog
// group keeps its own table next to the others; `locales/<lang>/commands.ts` flattens them.
export type Translations = Readonly<
	Record<string, readonly [fr: string, de: string, es: string, zhCN: string]>
>;

export type CommandLocale = 'en' | 'fr' | 'de' | 'es' | 'zh-CN';

const COLUMN: Readonly<Record<Exclude<CommandLocale, 'en'>, number>> = {
	fr: 0,
	de: 1,
	es: 2,
	'zh-CN': 3,
};

/** One locale's flat table (English is the identity map, as the locale coverage test expects). */
export function flatten(
	tables: readonly Translations[],
	locale: CommandLocale,
): Record<string, string> {
	const out: Record<string, string> = {};
	for (const table of tables)
		for (const [key, texts] of Object.entries(table))
			out[key] = locale === 'en' ? key : (texts[COLUMN[locale]] ?? key);
	return out;
}
