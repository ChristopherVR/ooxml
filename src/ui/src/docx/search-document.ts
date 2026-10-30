// The document as the text Find searches (`ooxml-core/docx/ui` matches it), and the transactions
// that apply a replacement to a match. A paragraph mark is `\r` only between two paragraphs of
// the same container, so a match never runs from one table cell into another.
import {
	Fragment,
	Slice,
	type Mark,
	type Node as ProseMirrorNode,
	type Schema,
} from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import {
	expectDefined,
	findMatches,
	FIND_LINE_BREAK,
	FIND_OBJECT,
	FIND_PAGE_BREAK,
	FIND_PARAGRAPH_MARK,
	insertedTextMarks,
	type FindOptions,
	type ReplacementPart,
} from 'ooxml-core/docx/ui';

export interface DocumentMatch {
	readonly from: number;
	readonly to: number;
	readonly text: string;
	readonly groups: readonly string[];
}

export interface ReplacementTarget {
	readonly from: number;
	readonly to: number;
	readonly parts: readonly ReplacementPart[];
}

interface FindableText {
	readonly text: string;
	/** The document position where each character of `text` starts. */
	readonly from: readonly number[];
	/** The document position where each character of `text` ends. */
	readonly to: readonly number[];
}

/** The document's text with each character's position; objects and breaks stand for their codes. */
export function findableText(doc: ProseMirrorNode): FindableText {
	let text = '';
	const from: number[] = [];
	const to: number[] = [];
	const add = (char: string, start: number, end: number): void => {
		text += char;
		from.push(start);
		to.push(end);
	};
	let previous: { parent: ProseMirrorNode; end: number } | null = null;
	const inline = (parent: ProseMirrorNode, start: number): void => {
		parent.forEach((child, offset) => {
			const pos = start + offset;
			if (child.isText) {
				const value = child.text ?? '';
				for (let index = 0; index < value.length; index++)
					add(value.charAt(index), pos + index, pos + index + 1);
			} else if (child.type.name === 'hardBreak') {
				add(FIND_LINE_BREAK, pos, pos + child.nodeSize);
			} else if (child.type.name === 'pageBreak' && child.attrs.kind !== 'column') {
				add(FIND_PAGE_BREAK, pos, pos + child.nodeSize);
			} else if (child.isLeaf) {
				add(FIND_OBJECT, pos, pos + child.nodeSize);
			} else {
				inline(child, pos + 1);
			}
		});
	};
	const block = (parent: ProseMirrorNode, start: number): void => {
		parent.forEach((child, offset) => {
			const pos = start + offset;
			if (child.isTextblock) {
				const contentStart = pos + 1;
				if (previous)
					add(
						previous.parent === parent ? FIND_PARAGRAPH_MARK : FIND_OBJECT,
						previous.end,
						contentStart,
					);
				inline(child, contentStart);
				previous = { parent, end: pos + child.nodeSize - 1 };
			} else if (child.isLeaf) {
				add(FIND_OBJECT, pos, pos + child.nodeSize);
				previous = null;
			} else {
				block(child, pos + 1);
			}
		});
	};
	block(doc, 0);
	return { text, from, to };
}

/** Every match of the query in the document, in document order. */
export function findDocumentMatches(
	doc: ProseMirrorNode,
	query: string,
	options: FindOptions = {},
): DocumentMatch[] {
	const flat = findableText(doc);
	return findMatches(flat.text, query, options).map((match) => ({
		from: expectDefined(flat.from[match.start], 'match start'),
		to: expectDefined(flat.to[match.end - 1], 'match end'),
		text: match.text,
		groups: match.groups,
	}));
}

/** The marks a replacement inherits: those of the text at the match, as the old search used. */
function marksAt(doc: ProseMirrorNode, pos: number): readonly Mark[] {
	const $pos = doc.resolve(pos);
	if ($pos.nodeAfter?.isText) return $pos.nodeAfter.marks;
	if ($pos.nodeBefore?.isText) return $pos.nodeBefore.marks;
	return $pos.marks();
}

/**
 * The paragraphs of one replacement. Inline content is gathered per paragraph; a paragraph mark
 * (`^p`, or a `\r` in matched text) starts the next paragraph, and a control character that is not
 * a break or a character (a `\0` object) is dropped.
 */
function replacementParagraphs(
	schema: Schema,
	parts: readonly ReplacementPart[],
	marks: readonly Mark[],
): ProseMirrorNode[][] {
	const paragraphs: ProseMirrorNode[][] = [[]];
	let buffer = '';
	const flush = (): void => {
		if (!buffer) return;
		const content = paragraphs[paragraphs.length - 1];
		content?.push(schema.text(buffer, insertedTextMarks(marks, buffer)));
		buffer = '';
	};
	const inline = (node: ProseMirrorNode): void => {
		flush();
		paragraphs[paragraphs.length - 1]?.push(node);
	};
	const text = (value: string): void => {
		for (const char of value) {
			if (char === FIND_PARAGRAPH_MARK) {
				flush();
				paragraphs.push([]);
			} else if (char === FIND_LINE_BREAK) {
				inline(
					expectDefined(schema.nodes.hardBreak, 'hardBreak node type').create(null, null, marks),
				);
			} else if (char === FIND_PAGE_BREAK) {
				inline(
					expectDefined(schema.nodes.pageBreak, 'pageBreak node type').create(
						{ kind: 'page' },
						null,
						marks,
					),
				);
			} else if (char !== FIND_OBJECT) {
				buffer += char;
			}
		}
	};
	for (const part of parts) {
		if (part.kind === 'text') text(part.text);
		else if (part.kind === 'paragraph') {
			flush();
			paragraphs.push([]);
		} else if (part.kind === 'lineBreak') text(FIND_LINE_BREAK);
		else if (part.kind === 'pageBreak') text(FIND_PAGE_BREAK);
		else if (part.kind === 'tab') text('\t');
		else if (part.kind === 'optionalHyphen') text('­');
		else if (part.kind === 'nonBreakingHyphen') text('‑');
		else text(' ');
	}
	flush();
	return paragraphs;
}

/**
 * One transaction that replaces every target. Targets are applied last first so the earlier
 * positions still hold. A replacement is spliced as an open slice of paragraphs, so a match that
 * spans paragraph marks joins them and a `^p` in the replacement splits them.
 */
export function replaceTargets(
	state: EditorState,
	targets: readonly ReplacementTarget[],
): Transaction {
	const { schema } = state;
	const ids = new Set<string>();
	state.doc.descendants((node) => {
		if (typeof node.attrs.id === 'string' && node.attrs.id) ids.add(node.attrs.id);
	});
	let counter = 0;
	const nextId = (): string => {
		let id = '';
		do id = `dve-find-${++counter}`;
		while (ids.has(id));
		ids.add(id);
		return id;
	};
	const paragraph = expectDefined(schema.nodes.paragraph, 'paragraph node type');
	let transaction = state.tr;
	for (const target of [...targets].sort((left, right) => right.from - left.from)) {
		const $from = state.doc.resolve(target.from);
		const attrs = $from.parent.attrs;
		const blocks = replacementParagraphs(schema, target.parts, marksAt(state.doc, target.from));
		const nodes = blocks.map((content, index) =>
			paragraph.create(index === 0 ? attrs : { ...attrs, id: nextId() }, content),
		);
		transaction = transaction.replace(
			target.from,
			target.to,
			new Slice(Fragment.from(nodes), 1, 1),
		);
	}
	return transaction;
}
