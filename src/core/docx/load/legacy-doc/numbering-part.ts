// word/numbering.xml from the .doc list tables: one abstract numbering per list definition and one
// numbering instance per list override (`ilfo`), with its level overrides.
import type { DocListLevel } from '@christophervr/ole2/doc-lists';
import type { Conversion } from './context';
import { paragraphPropertiesXml } from './para-props';
import { runPropertiesXml, wrapRunProperties } from './run-props';
import { el, ROOT_NAMESPACES, XML_DECLARATION } from './xml';

// Word's number format codes (nfc) and their ST_NumberFormat names.
const NUMBER_FORMATS = [
	'decimal',
	'upperRoman',
	'lowerRoman',
	'upperLetter',
	'lowerLetter',
	'ordinal',
	'cardinalText',
	'ordinalText',
	'hex',
	'chicago',
	'ideographDigital',
	'japaneseCounting',
	'aiueo',
	'iroha',
	'decimalFullWidth',
	'decimalHalfWidth',
	'japaneseLegal',
	'japaneseDigitalTenThousand',
	'decimalEnclosedCircle',
	'decimalFullWidth2',
	'aiueoFullWidth',
	'irohaFullWidth',
	'decimalZero',
	'bullet',
	'ganada',
	'chosung',
	'decimalEnclosedFullstop',
	'decimalEnclosedParen',
	'decimalEnclosedCircleChinese',
	'ideographEnclosedCircle',
	'ideographTraditional',
	'ideographZodiac',
	'ideographZodiacTraditional',
	'taiwaneseCounting',
	'ideographLegalTraditional',
	'taiwaneseCountingThousand',
	'taiwaneseDigital',
	'chineseCounting',
	'chineseLegalSimplified',
	'chineseCountingThousand',
	'koreanDigital',
	'koreanCounting',
	'koreanLegal',
	'koreanDigital2',
	'hebrew1',
	'arabicAlpha',
	'hebrew2',
	'arabicAbjad',
	'hindiVowels',
	'hindiConsonants',
	'hindiNumbers',
	'hindiCounting',
	'thaiLetters',
	'thaiNumbers',
	'thaiCounting',
	'vietnameseCounting',
	'numberInDash',
	'russianLower',
	'russianUpper',
];
const SUFFIX = { tab: undefined, space: 'space', nothing: 'nothing' } as const;
const hex8 = (value: number): string => (value >>> 0).toString(16).padStart(8, '0').toUpperCase();

function levelXml(level: DocListLevel, conversion: Conversion): string {
	const format = level.numberFormat === 255 ? 'none' : NUMBER_FORMATS[level.numberFormat];
	if (!format)
		conversion.warn(
			`A list level uses number format ${level.numberFormat}, which has no Word equivalent here; it is shown as decimal.`,
		);
	let inner = el('w:start', { 'w:val': level.startAt });
	inner += el('w:numFmt', { 'w:val': format ?? 'decimal' });
	if (level.noRestart) inner += el('w:lvlRestart', { 'w:val': 0 });
	if (level.legal) inner += el('w:isLgl');
	if (SUFFIX[level.follow]) inner += el('w:suff', { 'w:val': SUFFIX[level.follow] });
	inner += el('w:lvlText', { 'w:val': level.text });
	inner += el('w:lvlJc', { 'w:val': level.alignment });
	const pPr = paragraphPropertiesXml(level.paragraph, {}, { numbering: false });
	if (pPr) inner += `<w:pPr>${pPr}</w:pPr>`;
	const rPr = runPropertiesXml(level.character, { fonts: {} }, conversion, { style: false });
	inner += wrapRunProperties(rPr);
	return el('w:lvl', { 'w:ilvl': level.level }, inner);
}

export function numberingPart(conversion: Conversion): string | undefined {
	const { definitions, overrides } = conversion.model.lists;
	if (definitions.length === 0 && overrides.length === 0) return undefined;
	let body = '';
	definitions.forEach((definition, index) => {
		const levels = definition.levels.map((level) => levelXml(level, conversion)).join('');
		body += el(
			'w:abstractNum',
			{ 'w:abstractNumId': index },
			el('w:nsid', { 'w:val': hex8(definition.lsid) }) +
				el('w:multiLevelType', {
					'w:val': definition.simple
						? 'singleLevel'
						: definition.hybrid
							? 'hybridMultilevel'
							: 'multilevel',
				}) +
				el('w:tmpl', { 'w:val': hex8(definition.templateId) }) +
				levels,
		);
	});
	for (const override of overrides) {
		if (override.definitionIndex === undefined) continue;
		let inner = el('w:abstractNumId', { 'w:val': override.definitionIndex });
		for (const item of override.levels) {
			const start =
				item.startAt === undefined ? '' : el('w:startOverride', { 'w:val': item.startAt });
			const definition = item.definition ? levelXml(item.definition, conversion) : '';
			if (start || definition)
				inner += el('w:lvlOverride', { 'w:ilvl': item.level }, start + definition);
		}
		body += el('w:num', { 'w:numId': override.ilfo }, inner);
	}
	return `${XML_DECLARATION}<w:numbering ${ROOT_NAMESPACES}>${body}</w:numbering>`;
}
