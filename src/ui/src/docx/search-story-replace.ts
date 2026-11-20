// Replace and Replace All across stories: the body and an open story editor take a transaction of
// their own; a closed header, footer or note is rewritten from its model copy and handed to the
// story host, and text boxes are rewritten in the body. Each story is one undo step, with every
// footnote one story and every endnote another, as Word groups them.
import { closeHistory } from 'prosemirror-history';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Block, DocumentModel } from 'ooxml-core/docx';
import { expectDefined, replacementParts, type FindStory } from 'ooxml-core/docx/ui';
import { replaceTargets, type ReplacementTarget } from './search-document';
import { textBoxDrawingAttrs } from './text-box-commands';
import {
	replacedStoryBlocks,
	replacedTextBoxLines,
	type SearchUnit,
	type StoryMatch,
} from './search-stories';

export type HeaderFooterStory = Extract<FindStory, { kind: 'header' | 'footer' }>;

/** How Find reaches the stories outside the body; the editor's parts controller provides it. */
export interface SearchStoryHost {
	model(): DocumentModel;
	/** The open header, footer or note editor, keyed like `findStoryKey`. */
	openStory(): { key: string; view: EditorView } | undefined;
	/** Shows the story and selects `from`..`to` in its editor when it can open one. */
	showStory(story: FindStory, from: number, to: number): void;
	/** Closes an open story editor; does nothing when none is open. */
	closeStory(): void;
	/** Headers, footers and notes are locked in collaboration and forms protection. */
	canEditStories(): boolean;
	writeHeaderFooter(story: HeaderFooterStory, blocks: Block[]): void;
	writeNotes(kind: 'footnote' | 'endnote', notes: ReadonlyMap<string, Block[]>): void;
}

interface Entry {
	readonly match: StoryMatch;
	readonly target: ReplacementTarget;
}

/** One transaction over a view, the caret after the first replacement. */
function replaceInView(view: EditorView, entries: readonly Entry[]): number {
	const transaction = replaceTargets(
		view.state,
		entries.map((entry) => entry.target),
	);
	const first = expectDefined(entries[0], 'first match').match;
	const cursor = transaction.mapping.map(first.to, 1);
	const selection = TextSelection.near(transaction.doc.resolve(cursor));
	view.dispatch(closeHistory(transaction.setSelection(selection).scrollIntoView()));
	return entries.length;
}

/** Every editable text box with matches, rewritten in one body transaction. */
function replaceTextBoxes(view: EditorView, entries: readonly Entry[]): number {
	let transaction = view.state.tr;
	let count = 0;
	for (const [unit, own] of byUnit(entries)) {
		const pos = unit.textBoxPos;
		const node = pos === undefined ? null : transaction.doc.nodeAt(pos);
		if (pos === undefined || !node || !unit.replaceable) continue;
		const lines = replacedTextBoxLines(
			unit,
			own.map((entry) => entry.target),
		);
		transaction = transaction.setNodeMarkup(pos, undefined, {
			...node.attrs,
			textBoxText: JSON.stringify(lines),
			...textBoxDrawingAttrs(lines, node.attrs.textBoxBorder !== false, node.attrs),
		});
		count += own.length;
	}
	if (count) view.dispatch(closeHistory(transaction));
	return count;
}

function byUnit(entries: readonly Entry[]): Map<SearchUnit, Entry[]> {
	const units = new Map<SearchUnit, Entry[]>();
	for (const entry of entries) {
		const list = units.get(entry.match.unit) ?? [];
		list.push(entry);
		units.set(entry.match.unit, list);
	}
	return units;
}

/**
 * Applies the replacement to `matches` (in search order) and returns how many were replaced. A
 * story that cannot change (read-only text boxes, locked headers and notes) is skipped. Throws
 * `FindPatternError` before changing anything when the replacement is invalid.
 */
export function applyReplacements(
	view: EditorView,
	host: SearchStoryHost | undefined,
	matches: readonly StoryMatch[],
	options: { readonly replacement: string; readonly wildcards: boolean },
): number {
	const entries: Entry[] = matches.map((match) => ({
		match,
		target: {
			from: match.from,
			to: match.to,
			parts: replacementParts(options.replacement, match, options.wildcards),
		},
	}));
	const groups = new Map<string, Entry[]>();
	for (const entry of entries) {
		const list = groups.get(entry.match.unit.group) ?? [];
		list.push(entry);
		groups.set(entry.match.unit.group, list);
	}
	let count = 0;
	for (const [group, own] of groups) {
		const { unit } = expectDefined(own[0], 'group match').match;
		if (group === 'main') {
			count += replaceInView(view, own);
			continue;
		}
		if (!host) continue;
		if (group === 'textBoxes') {
			count += replaceTextBoxes(view, own);
			continue;
		}
		if (!host.canEditStories()) continue;
		const open = host.openStory();
		if (open && own.every((entry) => entry.match.unit.key === open.key)) {
			count += replaceInView(open.view, own);
			continue;
		}
		const { story } = unit;
		if (story.kind === 'footnote' || story.kind === 'endnote') {
			const notes = new Map<string, Block[]>();
			for (const [noteUnit, list] of byUnit(own))
				if (noteUnit.story.kind === 'footnote' || noteUnit.story.kind === 'endnote')
					notes.set(
						noteUnit.story.id,
						replacedStoryBlocks(
							noteUnit,
							list.map((entry) => entry.target),
						),
					);
			host.writeNotes(story.kind, notes);
		} else if (story.kind === 'header' || story.kind === 'footer')
			host.writeHeaderFooter(
				story,
				replacedStoryBlocks(
					unit,
					own.map((entry) => entry.target),
				),
			);
		else continue;
		count += own.length;
	}
	return count;
}
