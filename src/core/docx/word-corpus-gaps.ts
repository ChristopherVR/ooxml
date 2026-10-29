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
		id: 'content-control-line-height',
		name: '11-content-controls',
		paragraphs: [5],
		aspect: 'y',
		reason:
			'The fifth paragraph starts 2.0 pt higher than in Word, after a paragraph holding a ' +
			'check-box content control; the cause is not isolated.',
	},
];
