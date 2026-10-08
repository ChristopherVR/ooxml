import type { VisioTextEdit } from '../edit-commands';

/** UTF-16 bounds; oversized inputs fail instead of producing partial replacements. */
export const VISIO_TEXT_REPLACE_LIMITS = Object.freeze({
	query: 256,
	replacement: 32768,
	text: 2_000_000,
	names: 1_000_000,
	output: 1_000_000,
	matches: 10000,
	commands: 1000,
	shapes: 25000,
	pages: 256,
	depth: 64,
});
export interface VisioTextReplaceScope {
	query: string;
	/** Literal case-sensitive matching. Unicode case folding is not supported. */
	matchCase: true;
	scope: 'selection' | 'current-page' | 'all-pages';
	pageId: string;
	selection?: readonly { id: string; pageId?: string }[];
}
export interface VisioTextOccurrence {
	readonly pageId: string;
	readonly shapeId: string;
	/** Zero-based UTF-16 range, excluding end. */
	readonly start: number;
	readonly end: number;
}
export interface VisioTextReplaceRequest extends VisioTextReplaceScope {
	replacement: string;
	mode: 'current' | 'all';
	/** Omitted selects the first occurrence. Supplied ranges must still match exactly. */
	current?: VisioTextOccurrence;
}
export interface VisioTextReplacePlan {
	readonly occurrences: readonly VisioTextOccurrence[];
	readonly replacementCount: number;
	readonly edits: readonly Readonly<VisioTextEdit>[];
	readonly nextOccurrence: VisioTextOccurrence | undefined;
}
export const textReplaceError = (message: string): never => {
	throw new Error(message);
};
export function replaceString(value: unknown, maximum: number, label: string): string {
	if (typeof value !== 'string' || value.length > maximum)
		return textReplaceError(`Replace ${label} exceeds supported string limits.`);
	return value;
}
export function replaceIdentity(value: unknown): string {
	const id = replaceString(value, 256, 'identity');
	if (!id) textReplaceError('Replace requires nonempty identities.');
	return id;
}
