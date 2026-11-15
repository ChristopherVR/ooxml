import type { SearchStoryInfo } from './search-controller';
import { translate, translateTemplate, type EditorLocale } from './localization';

/** Find codes for the Special menu, in Word's order: the English label and the code it inserts. */
export const SPECIAL_CODES: ReadonlyArray<readonly [label: string, code: string]> = [
	['Paragraph Mark', '^p'],
	['Tab Character', '^t'],
	['Any Character', '^?'],
	['Any Digit', '^#'],
	['Any Letter', '^$'],
	['Caret Character', '^^'],
	['Manual Line Break', '^l'],
	['Page Break', '^m'],
	['Optional Hyphen', '^-'],
	['Nonbreaking Hyphen', '^~'],
	['Nonbreaking Space', '^s'],
	['White Space', '^w'],
];

/** The Find options of More >>, each with its English label. */
export const OPTION_LABELS = [
	['wholeWords', 'Find whole words only'],
	['wildcards', 'Use wildcards'],
	['matchPrefix', 'Match prefix'],
	['matchSuffix', 'Match suffix'],
	['ignorePunctuation', 'Ignore punctuation characters'],
	['ignoreWhitespace', 'Ignore white-space characters'],
] as const;

/** Word disables these options while Use wildcards is on. */
export const WILDCARD_DISABLED = new Set<string>([
	'wholeWords',
	'matchPrefix',
	'matchSuffix',
	'ignorePunctuation',
	'ignoreWhitespace',
]);

const HEADER_FOOTER_NAMES = {
	header: { default: 'Header', first: 'First page header', even: 'Even page header' },
	footer: { default: 'Footer', first: 'First page footer', even: 'Even page footer' },
} as const;

/** The story a match is in, as Word names it: "Footnotes", "First page header -Section 2-". */
export function storyLabel(locale: EditorLocale, story: SearchStoryInfo | undefined): string {
	if (!story) return '';
	switch (story.kind) {
		case 'main':
			return translate(locale, 'Main Document');
		case 'footnote':
			return translate(locale, 'Footnotes');
		case 'endnote':
			return translate(locale, 'Endnotes');
		case 'textBoxes':
			return translate(locale, 'Text Boxes');
		default:
			return `${translate(locale, HEADER_FOOTER_NAMES[story.kind][story.slot ?? 'default'])} ${translateTemplate(
				locale,
				'-Section {section}-',
				{ section: (story.sectionIndex ?? 0) + 1 },
			)}`;
	}
}
