// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
/**
 * Word's outline numbering for levels 1 to 9, as prefix, `w:numFmt` and suffix: I. A. 1. a) (1)
 * (a) (i) (a) (i). AUTONUMOUT numbers headings with it and LISTNUM's OutlineDefault list draws it.
 */
export const OUTLINE_NUMBERING: readonly [string, string, string][] = [
	['', 'upperRoman', '.'],
	['', 'upperLetter', '.'],
	['', 'decimal', '.'],
	['', 'lowerLetter', ')'],
	['(', 'decimal', ')'],
	['(', 'lowerLetter', ')'],
	['(', 'lowerRoman', ')'],
	['(', 'lowerLetter', ')'],
	['(', 'lowerRoman', ')'],
];
