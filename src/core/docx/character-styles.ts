// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { CharacterStyleDefinition, RunStyleCatalog } from './run-style-model';
import { first, getW, named, parseXml, type XmlDocument, type XmlElement } from './xml';
import { parseRunProperties } from './run-properties';
import { parseOnOff } from './simple-types';
import { inheritanceCycleEntries } from './style-inheritance';

function styleElements(document: XmlDocument): XmlElement[] {
	return Array.from(document.getElementsByTagName('*')).filter(
		(element): element is XmlElement =>
			named(element, 'style') &&
			(getW(element, 'type') === 'character' || getW(element, 'type') === 'paragraph'),
	);
}

const eastAsiaLanguageOf = (rPr: XmlElement | undefined) => getW(first(rPr, 'lang'), 'eastAsia');

/** Parses `styles.xml`'s `rPrDefault` plus every character/paragraph style's `w:rPr`. */
export function parseRunStyleCatalog(xml: string): RunStyleCatalog {
	const document = parseXml(xml);
	const defaults = first(first(document.documentElement, 'docDefaults'), 'rPrDefault');
	const styles = Object.create(null) as Record<string, CharacterStyleDefinition>;
	for (const element of styleElements(document)) {
		const id = getW(element, 'styleId');
		if (!id) continue;
		const type = getW(element, 'type') === 'character' ? 'character' : 'paragraph';
		const basedOn = getW(first(element, 'basedOn'), 'val');
		const linkedStyle = getW(first(element, 'link'), 'val');
		const isDefault = parseOnOff(getW(element, 'default'));
		const name = getW(first(element, 'name'), 'val');
		const eastAsiaLanguage = eastAsiaLanguageOf(first(element, 'rPr'));
		styles[id] = {
			id,
			type,
			...(name ? { name } : {}),
			...(basedOn ? { basedOn } : {}),
			...(linkedStyle ? { linkedStyle } : {}),
			...(isDefault === undefined ? {} : { isDefault }),
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
		warnings,
	};
}
