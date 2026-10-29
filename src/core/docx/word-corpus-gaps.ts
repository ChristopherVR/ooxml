// Known differences between the core's layout and the Word reference corpus (see
// `word-corpus.test.ts`). A gap names a document, the Word paragraph numbers it covers and the
// aspect that differs (`y`, `x` or `lines`). The test fails when a listed paragraph starts to
// match, so an entry must be deleted as soon as its cause is fixed.
export interface AcceptedGap {
	id: string;
	name: string;
	/** Word paragraph numbers (1-based, as in the facts file). */
	paragraphs: number[];
	aspect: 'y' | 'x' | 'lines';
	reason: string;
}

export const acceptedGaps: AcceptedGap[] = [
	{
		id: 'bullet-marker-line-height',
		name: '02-lists-numbering',
		paragraphs: [3, 4, 5, 6, 7, 8, 9, 10],
		aspect: 'y',
		reason:
			'Word spaces the three bullet lines 14.2 pt apart where a text line is 12.5 pt (the ' +
			'bullet glyph font is taller); the layout ignores the marker font in line height. Each ' +
			'bullet adds about 1.8 pt, so every later paragraph is 2.2 to 6.1 pt high (the drift grows with each bullet ' +
			'above it).',
	},
	{
		id: 'indented-first-character-x',
		name: '01-headings-styles',
		paragraphs: [5, 7],
		aspect: 'x',
		reason:
			"Word's Range.Information(5) reports 72 pt (the column edge) for the first character of an " +
			'indented paragraph (36 pt left, 18 pt first line: the layout starts text at 126 pt) and ' +
			'of a centred one (84.9 pt here). The recorder reads the position of a collapsed range at ' +
			'the paragraph start; whether that is the column edge or the true caret position was not ' +
			'established, so this is a facts-mapping question and not evidence of a layout error.',
	},
	{
		id: 'table-height',
		name: '03-tables-merged',
		paragraphs: [20],
		aspect: 'y',
		reason:
			'The paragraph after the merged-cell table starts 28.7 pt lower than in Word: the ' +
			'layout’s table height differs from Word’s. Table layout is being reworked ' +
			'separately, so this is recorded rather than adjusted here.',
	},
	{
		id: 'section-break-spacing',
		name: '06-sections-columns',
		paragraphs: [5, 6, 9, 10],
		aspect: 'y',
		reason:
			'The first heading of a section that starts on a new page is 6 pt lower than Word: Word ' +
			'collapses its space before (18 pt) against the space after (6 pt) of the section-break ' +
			'paragraph that ends the previous section, the layout starts each section without that ' +
			'neighbour. The paragraph after it keeps the 6 pt offset.',
	},
	{
		id: 'display-equation-height',
		name: '08-equation',
		paragraphs: [5],
		aspect: 'y',
		reason:
			'The paragraph after a display equation is 6.8 pt lower than in Word: the height of the ' +
			'equation block differs.',
	},
	{
		id: 'content-control-line-height',
		name: '11-content-controls',
		paragraphs: [5],
		aspect: 'y',
		reason:
			'The fifth paragraph starts 2.0 pt higher than in Word, after a paragraph holding a ' +
			'check-box content control; the cause is not isolated.',
	},
];
