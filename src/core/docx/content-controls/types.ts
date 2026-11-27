// The public shapes of content control templating (`fillContentControls`, `readContentControls`,
// `flattenContentControls`): see docs/docx-content-controls.md.
import type { ContentControlLock } from '../content-control-model';
import type { HtmlMergeIssue } from '../mail-merge/html-format';
import type { HeaderFooterSlot, MergeStory } from '../mail-merge/stories';
import type { DocumentModel } from '../model';

export interface ContentControlFillData {
	/** Values by control tag (`w:tag`). */
	values: Readonly<Record<string, string>>;
	/** The tags whose value is HTML; any other value is plain text, whatever it holds. */
	html?: ReadonlyArray<string>;
	/** By tag of the control (or of the control in a table row) to repeat: one record per copy. */
	repeat?: Readonly<Record<string, ReadonlyArray<Readonly<Record<string, string>>>>>;
}

export interface ContentControlFillOptions {
	/** A control whose tag ends with it is removed when it, and every control inside it, is empty. */
	optionalSuffix?: string;
	/** Locks to set by tag after filling (`none` removes the lock). */
	lock?: Readonly<Record<string, ContentControlLock | 'none'>>;
	/** When filling an already filled document: the template its structure comes from. */
	template?: DocumentModel;
}

export type ContentControlFillStatus =
	/** The control holds the value now. */
	| 'filled'
	/** The value was empty: the control holds one empty paragraph or run. */
	| 'cleared'
	/** The control was copied once per item (`items`). */
	| 'repeated'
	/** An optional control (or a repeated one without items) was removed with its content. */
	| 'removed'
	/** An optional control stays: it, or a control inside it, is not empty. */
	| 'kept'
	/** Its lock was set. */
	| 'locked'
	/** Nothing was written: `reason` says why. */
	| 'skipped'
	/** No control has the tag named in the data. */
	| 'missing';

export interface ContentControlFillResult {
	tag: string;
	status: ContentControlFillStatus;
	story?: MergeStory;
	/** In a header or footer, the section slot. */
	slot?: HeaderFooterSlot;
	/** Why nothing was written, or what was written instead (a value read as text). */
	reason?: string;
	/** A repeated control's number of copies. */
	items?: number;
	/** Inside a repeated control, the copy (1 for the first). */
	item?: number;
}

export interface FillContentControlsResult {
	model: DocumentModel;
	results: ContentControlFillResult[];
	/** What an HTML value lost on the way in (refused pictures and links, flattened tables). */
	htmlIssues: HtmlMergeIssue[];
}

export interface ReadContentControlsOptions {
	/** Only tags that start with it. */
	prefix?: string;
	/** `text` (default) or `html`. */
	as?: 'text' | 'html';
}

export interface FlattenContentControlsOptions {
	/** Also remove the paragraphs whose paragraph mark and content are all hidden text. */
	removeHiddenParagraphs?: boolean;
}
