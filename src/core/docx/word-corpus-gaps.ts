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

export const acceptedGaps: AcceptedGap[] = [];
