import { readStyleMetadata } from './style-metadata';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { CharacterStyleDefinition, RunStyleCatalog } from './run-style-model';
import { children, first, getW, parseXml, type XmlElement } from './xml';
import { parseRunProperties } from './run-properties';
import { inheritanceCycleEntries } from './style-inheritance';

const eastAsiaLanguageOf = (rPr: XmlElement | undefined) => getW(first(rPr, 'lang'), 'eastAsia');

/** Parses `styles.xml`'s `rPrDefault` plus every character/paragraph style's `w:rPr`. */
export function parseRunStyleCatalog(xml: string | XmlElement): RunStyleCatalog {
	const root = typeof xml === 'string' ? parseXml(xml).documentElement : xml;
	const defaults = first(first(root, 'docDefaults'), 'rPrDefault');
	const styles = Object.create(null) as Record<string, CharacterStyleDefinition>;
	const reservedStyleIds: string[] = [];
	for (const element of children(root, 'style')) {
		const type = getW(element, 'type');
		const id = getW(element, 'styleId');
		if (!id) continue;
		if (type !== 'character' && type !== 'paragraph') {
			if (type !== 'table') reservedStyleIds.push(id);
			continue;
		}
		const linkedStyle = getW(first(element, 'link'), 'val');
		const eastAsiaLanguage = eastAsiaLanguageOf(first(element, 'rPr'));
		styles[id] = {
			...readStyleMetadata(element, id),
			type,
			...(linkedStyle ? { linkedStyle } : {}),
			formatting: parseRunProperties(first(element, 'rPr')),
			...(eastAsiaLanguage ? { eastAsiaLanguage } : {}),
		};
	}
	const warnings = inheritanceCycleEntries(styles).map(
		(id) => `Character style inheritance cycle detected at "${id}"; cyclic inheritance is ignored.`,
	);
	const docDefaultsEastAsiaLanguage = eastAsiaLanguageOf(first(defaults, 'rPr'));
	return {
		docDefaults: parseRunProperties(first(defaults, 'rPr')),
		...(docDefaultsEastAsiaLanguage ? { docDefaultsEastAsiaLanguage } : {}),
		styles,
		...(reservedStyleIds.length && { reservedStyleIds }),
		warnings,
	};
}
