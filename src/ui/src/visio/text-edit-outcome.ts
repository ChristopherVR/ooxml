/**
 * The result of committing the canvas's inline text draft (`commitTextEdit`). Every outcome names
 * the source and document generation it was decided in, and the page and shape when a draft
 * existed, so a host can tell which document an answer belongs to.
 */
export interface TextEditIdentity {
	/** The loaded source's generation (the draft's own when there was one). */
	readonly sourceGeneration: number;
	/** The accepted-document generation when the outcome was decided. */
	readonly documentGeneration: number;
	readonly pageId?: string;
	readonly shapeId?: string;
}

export type TextEditOutcome =
	/** The draft was applied as one undoable edit. */
	| (TextEditIdentity & { readonly status: 'committed'; readonly text: string })
	/** Nothing to apply: no draft was open, or the draft equals the shape's text. */
	| (TextEditIdentity & { readonly status: 'unchanged'; readonly reason: 'no-draft' | 'same-text' })
	/** The editor refused the draft; `text` is the rejected draft, kept for recovery. */
	| (TextEditIdentity & {
			readonly status: 'refused';
			readonly reason: 'read-only' | 'rejected';
			readonly message: string;
			readonly text: string;
	  })
	/**
	 * The draft was not applied because its document went away or the edit was superseded, so it
	 * never reaches another document. `text` is the draft when there was one.
	 */
	| (TextEditIdentity & {
			readonly status: 'cancelled';
			readonly reason: 'stale-document' | 'superseded' | 'destroyed';
			readonly text?: string;
	  });

export type TextEditCommitter = () => Promise<TextEditOutcome>;
