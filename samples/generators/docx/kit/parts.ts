/**
 * The parts `saveDocx` does not write for a new document, written by hand: `word/numbering.xml`
 * (the standalone writer refuses numbering definitions) and `word/settings.xml` (the core only
 * patches single settings into it).
 */
import { W_NS } from './builder';
import { paragraphLook, runLook, type ParagraphLook, type RunLook } from './styles';

export interface LevelSpec {
	/** `decimal`, `lowerLetter`, `lowerRoman`, `bullet`, `none`... */
	format: string;
	/** `%1.%2`, or the bullet character. */
	text: string;
	start?: number;
	/** The paragraph style bound to this level (heading numbering). */
	style?: string;
	suffix?: 'tab' | 'space' | 'nothing';
	paragraph?: ParagraphLook;
	run?: RunLook;
	/** Show every level number as arabic, whatever its own format (legal numbering). */
	legal?: boolean;
}

export interface ListSpec {
	/** `w:abstractNumId`; the `w:num` that uses it gets the same id plus one. */
	id: number;
	/** `multilevel`, `hybridMultilevel` or `singleLevel`. */
	type?: string;
	levels: LevelSpec[];
}

/** `numId` of the list instance `numberingXml` writes for `spec`. */
export const numIdOf = (spec: ListSpec): number => spec.id + 1;

export function numberingXml(lists: readonly ListSpec[]): string {
	const abstract = lists.map(
		(list) =>
			`<w:abstractNum w:abstractNumId="${list.id}"><w:multiLevelType w:val="${
				list.type ?? 'multilevel'
			}"/>${list.levels
				.map(
					(level, index) =>
						`<w:lvl w:ilvl="${index}"><w:start w:val="${level.start ?? 1}"/><w:numFmt w:val="${level.format}"/>${
							level.style ? `<w:pStyle w:val="${level.style}"/>` : ''
						}${level.legal ? '<w:isLgl/>' : ''}${
							level.suffix ? `<w:suff w:val="${level.suffix}"/>` : ''
						}<w:lvlText w:val="${level.text}"/><w:lvlJc w:val="left"/>${
							level.paragraph ? `<w:pPr>${paragraphLook(level.paragraph)}</w:pPr>` : ''
						}${level.run ? `<w:rPr>${runLook(level.run)}</w:rPr>` : ''}</w:lvl>`,
				)
				.join('')}</w:abstractNum>`,
	);
	const instances = lists.map(
		(list) => `<w:num w:numId="${numIdOf(list)}"><w:abstractNumId w:val="${list.id}"/></w:num>`,
	);
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W_NS}">${abstract.join('')}${instances.join('')}</w:numbering>`;
}

export interface SettingsSpec {
	/** Notes exist: the separators the core wrote are declared (`w:footnotePr`, `w:endnotePr`). */
	footnotes?: boolean;
	endnotes?: boolean;
	/** `w:numFmt` of endnote marks; Word's default is lower-case roman. */
	endnoteFormat?: string;
	evenAndOddHeaders?: boolean;
	autoHyphenation?: boolean;
	defaultTabStop?: number;
	/** `w:documentProtection w:edit`, unenforced unless `enforced`. */
	protection?: { edit: 'forms' | 'readOnly' | 'comments' | 'trackedChanges'; enforced: boolean };
	/** The language of theme fonts and of the document, such as `en-GB`. */
	language: string;
	/** Maths is present: paragraph defaults for display equations. */
	math?: boolean;
}

/** `word/settings.xml` in CT_Settings order, Word 2013+ layout (compatibility mode 15). */
export function settingsXml(spec: SettingsSpec): string {
	const compat = (name: string, value: string) =>
		`<w:compatSetting w:name="${name}" w:uri="http://schemas.microsoft.com/office/word" w:val="${value}"/>`;
	return [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
		`<w:settings xmlns:w="${W_NS}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">`,
		'<w:zoom w:percent="100"/>',
		spec.protection
			? `<w:documentProtection w:edit="${spec.protection.edit}" w:enforcement="${spec.protection.enforced ? 1 : 0}"/>`
			: '',
		`<w:defaultTabStop w:val="${spec.defaultTabStop ?? 720}"/>`,
		spec.autoHyphenation ? '<w:autoHyphenation/>' : '',
		spec.evenAndOddHeaders ? '<w:evenAndOddHeaders/>' : '',
		'<w:characterSpacingControl w:val="doNotCompress"/>',
		spec.footnotes
			? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
			: '',
		spec.endnotes
			? `<w:endnotePr>${
					spec.endnoteFormat ? `<w:numFmt w:val="${spec.endnoteFormat}"/>` : ''
				}<w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>`
			: '',
		'<w:compat>',
		compat('compatibilityMode', '15'),
		compat('overrideTableStyleFontSizeAndJustification', '1'),
		compat('enableOpenTypeFeatures', '1'),
		compat('doNotFlipMirrorIndents', '1'),
		compat('differentiateMultirowTableHeaders', '1'),
		'</w:compat>',
		// No `m:mathFont`: the samples name no maths font, Word uses its own.
		spec.math
			? '<m:mathPr><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>'
			: '',
		`<w:themeFontLang w:val="${spec.language}"/>`,
		'<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>',
		'<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>',
		'</w:settings>',
	].join('');
}
