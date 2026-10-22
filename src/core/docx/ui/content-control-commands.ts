// Editing commands for content controls: toggling, picking, typing over a placeholder, inserting.
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import {
	afterEdit,
	canDeleteControl,
	canEditControlContent,
	checkboxState,
	placeholderTextOf,
	removedByEdit,
	selectListItem,
	setControlDate,
	toggleCheckbox,
	withValue,
	type ControlChange,
} from '../content-control-edit';
import {
	freshControlId,
	type ContentControl,
	type ContentControlKind,
} from '../content-control-model';
import {
	controlAt,
	controlSpans,
	everyControlSpan,
	type ControlSpan,
} from './content-control-spans';

/** Transactions carrying this meta are the control's own edits and bypass the lock guard. */
export const CONTROL_EDIT_META = 'dve-content-control';

const PLACEHOLDER_STYLE = 'PlaceholderText';
const isPlaceholderStyle = (mark: Mark) =>
	mark.type.name === 'characterStyle' && mark.attrs.id === PLACEHOLDER_STYLE;

/**
 * The first text of a control: its formatting carries over to the content that replaces it. A
 * cell's text is in its paragraphs, not in a table in it, which the replacement leaves as it is.
 */
function firstText(state: EditorState, span: ControlSpan): ProseMirrorNode | null {
	if (span.level === 'inline') return state.doc.nodeAt(span.from);
	let found: ProseMirrorNode | null = null;
	state.doc.nodesBetween(span.from, span.to, (node, pos) => {
		if (!found && node.isText) found = node;
		return !found && !(span.level === 'cell' && node.type.name === 'table' && pos > span.from);
	});
	return found;
}

/** Whether `doc` holds no inline content (text, a break, a picture) between `from` and `to`. */
function holdsNothing(doc: ProseMirrorNode, from: number, to: number): boolean {
	let nothing = true;
	if (from < to)
		doc.nodesBetween(from, to, (node) => {
			if (node.isInline) nothing = false;
			return nothing;
		});
	return nothing;
}

/** Whether a table in the cell at `cell` holds part of `from` to `to`. */
function touchesTable(doc: ProseMirrorNode, cell: number, from: number, to: number): boolean {
	let found = false;
	doc.nodesBetween(from, to, (node, pos) => {
		found ||= node.type.name === 'table' && pos > cell;
		return !found;
	});
	return found;
}

/**
 * Where the text of the cell at `pos` runs: from inside its first paragraph to the end of its
 * last. A table can come before or after them.
 */
function cellText(doc: ProseMirrorNode, pos: number): { from: number; to: number } {
	const cell = doc.nodeAt(pos)!;
	const text = { from: pos + 2, to: pos + cell.nodeSize - 2 };
	let first = true;
	cell.forEach((child, offset) => {
		if (child.type.name !== 'paragraph') return;
		if (first) text.from = pos + 2 + offset;
		text.to = pos + offset + child.nodeSize;
		first = false;
	});
	return text;
}

/**
 * The content of a cell control's cell once it holds `text`. A table holding a cell control keeps
 * its structure when it is saved, so the cell keeps its paragraphs and the tables in it: its first
 * paragraph holds the text and the others are emptied. Each keeps its formatting, not a control
 * around it: that goes with its content.
 */
function cellContent(
	state: EditorState,
	cell: ProseMirrorNode,
	text: ProseMirrorNode | null,
): Fragment {
	const nodes: ProseMirrorNode[] = [];
	let placed = false;
	cell.forEach((child) => {
		if (child.type.name !== 'paragraph') {
			nodes.push(child);
			return;
		}
		const attrs = child.attrs.contentControl
			? { ...child.attrs, contentControl: null }
			: child.attrs;
		nodes.push(child.type.create(attrs, placed ? null : text));
		placed = true;
	});
	if (!placed) nodes.push(state.schema.nodes.paragraph!.create(null, text));
	return Fragment.fromArray(nodes);
}

function contentMarks(state: EditorState, span: ControlSpan, change: ControlChange): Mark[] {
	const schema = state.schema;
	const first = firstText(state, span);
	// A placeholder is drawn in the Placeholder Text style, as Word writes it; a value is not.
	const placeholder = Boolean(change.control.showingPlaceholder && !change.removed);
	const kept = (first?.marks ?? []).filter(
		(mark) =>
			mark.type.name !== 'contentControl' &&
			!isPlaceholderStyle(mark) &&
			!(placeholder && mark.type.name === 'characterStyle'),
	);
	const marks = change.font
		? kept.filter((mark) => mark.type.name !== 'font')
		: kept.filter(() => true);
	if (change.font) {
		const font = kept.find((mark) => mark.type.name === 'font');
		marks.push(
			schema.marks.font!.create({ size: null, color: null, ...font?.attrs, family: change.font }),
		);
	}
	if (placeholder && schema.marks.characterStyle)
		marks.push(schema.marks.characterStyle.create({ id: PLACEHOLDER_STYLE }));
	if (span.level === 'inline' && !change.removed)
		marks.push(schema.marks.contentControl!.create({ data: JSON.stringify(change.control) }));
	return marks;
}

const placeholderChange = (control: ContentControl): ControlChange => ({
	control: { ...control, showingPlaceholder: true },
	text: placeholderTextOf(control),
});

/**
 * Replaces a control's content (and its state) as one edit. An inline control's text is replaced;
 * a cell control's cell holds the new text in its first paragraph (`cellContent`).
 */
export function replaceControlContent(
	state: EditorState,
	span: ControlSpan,
	edit: ControlChange,
): Transaction {
	// An inline control lives on its text, so an entry or a date that shows none (an empty display
	// text or date format) would take the control with it: it shows its placeholder instead, as a
	// control does once its text is deleted.
	const change =
		span.level === 'inline' && !edit.text && !edit.removed ? placeholderChange(edit.control) : edit;
	const marks = contentMarks(state, span, change);
	const text = change.text ? state.schema.text(change.text, marks) : null;
	if (span.level !== 'cell')
		return state.tr.replaceWith(span.from, span.to, text ?? []).setMeta(CONTROL_EDIT_META, true);
	const cell = state.doc.nodeAt(span.from)!;
	const tr = state.tr.replaceWith(span.from + 1, span.to - 1, cellContent(state, cell, text));
	tr.setNodeMarkup(span.from, undefined, {
		...cell.attrs,
		contentControl: change.removed ? null : change.control,
	});
	// A caret in the cell stays in it: mapped past the replaced content it would land in the next
	// cell. It goes to the start of a placeholder (typing replaces it), else after the text.
	const { from, to } = state.selection;
	if (from > span.from && to < span.to) {
		const shown = change.control.showingPlaceholder && !change.removed;
		const start = cellText(tr.doc, span.from).from;
		tr.setSelection(TextSelection.create(tr.doc, shown ? start : start + (text?.nodeSize ?? 0)));
	}
	return tr.setMeta(CONTROL_EDIT_META, true);
}

/**
 * Whether the content of `span` may be replaced as a whole: not when it holds a control that may
 * not be deleted (`sdtLocked`), one around a table or a row of a table in it included, since the
 * replacement carries `CONTROL_EDIT_META` past the lock guard.
 */
function replaceable(state: EditorState, span: ControlSpan): boolean {
	return !everyControlSpan(state.doc).some(
		(inner) =>
			inner.from >= span.from &&
			inner.to <= span.to &&
			(inner.level !== span.level || inner.from !== span.from) &&
			!canDeleteControl(inner.control),
	);
}

/**
 * Whether the content of `span` may be edited: by its own lock, and by those of the controls around
 * it (a row, cell or block control whose content is locked, one around a whole table included),
 * since the commands go past the lock guard.
 * A group control is the exception: the controls in a group follow their own locks.
 */
function editable(state: EditorState, span: ControlSpan): boolean {
	if (!canEditControlContent(span.control)) return false;
	return !everyControlSpan(state.doc).some(
		(around) =>
			around.control.kind !== 'group' &&
			!canEditControlContent(around.control) &&
			around.from <= span.from &&
			around.to >= span.to &&
			!(around.level === span.level && around.from === span.from && around.to === span.to),
	);
}

/** Stores new properties on every run or paragraph of the control. */
export function updateControlProperties(
	state: EditorState,
	span: ControlSpan,
	control: ContentControl,
): Transaction {
	const tr = state.tr.setMeta(CONTROL_EDIT_META, true);
	if (span.level === 'cell') {
		const cell = state.doc.nodeAt(span.from)!;
		return tr.setNodeMarkup(span.from, undefined, { ...cell.attrs, contentControl: control });
	}
	if (span.level === 'block') {
		state.doc.nodesBetween(span.from, span.to, (node, pos) => {
			if (node.type.name === 'paragraph' && node.attrs.contentControl?.id === span.control.id)
				tr.setNodeMarkup(pos, undefined, { ...node.attrs, contentControl: control });
			return node.type.name !== 'paragraph';
		});
		return tr;
	}
	const type = state.schema.marks.contentControl!;
	tr.removeMark(span.from, span.to, type);
	return tr.addMark(span.from, span.to, type.create({ data: JSON.stringify(control) }));
}

/**
 * The cell or inline control `id`. Ids can repeat (in files from other tools), so `at`, a position
 * in the control that was operated, picks the one meant: the innermost around it.
 */
function spanById(state: EditorState, id: string, at?: number): ControlSpan | undefined {
	const spans = controlSpans(state.doc).filter(
		(span) => span.level !== 'block' && span.control.id === id,
	);
	if (at === undefined || spans.length < 2) return spans[0];
	return spans
		.filter((span) => span.from <= at && at <= span.to)
		.sort((a, b) => a.to - a.from - (b.to - b.from))[0];
}

export function toggleCheckboxControl(
	state: EditorState,
	id: string,
	at?: number,
): Transaction | undefined {
	const span = spanById(state, id, at);
	if (
		!span ||
		span.control.kind !== 'checkbox' ||
		!editable(state, span) ||
		!replaceable(state, span)
	)
		return undefined;
	return replaceControlContent(state, span, toggleCheckbox(span.control));
}

export function chooseListItem(
	state: EditorState,
	id: string,
	index: number,
	at?: number,
): Transaction | undefined {
	const span = spanById(state, id, at);
	if (!span || !editable(state, span) || !replaceable(state, span)) return undefined;
	const change = selectListItem(span.control, index);
	return change ? replaceControlContent(state, span, change) : undefined;
}

export function chooseDate(
	state: EditorState,
	id: string,
	date: Date,
	at?: number,
): Transaction | undefined {
	const span = spanById(state, id, at);
	if (!span || span.control.kind !== 'date' || !editable(state, span) || !replaceable(state, span))
		return undefined;
	return replaceControlContent(state, span, setControlDate(span.control, date));
}

/**
 * The control whose rules hold for what is typed, pasted or deleted at `pos`: the innermost one,
 * except that a block control in a cell control (`w:tc > w:sdt > w:p`) leaves them to the cell.
 */
function inputControlAt(spans: readonly ControlSpan[], pos: number): ControlSpan | undefined {
	const span = controlAt(spans, pos, true);
	if (span?.level !== 'block') return span;
	const cells = spans.filter((other) => other.level === 'cell');
	return controlAt(cells, pos) ?? span;
}

/**
 * Typing over a control that shows its placeholder replaces the whole placeholder. Returns
 * `'blocked'` for kinds that take no typed text (checkbox, drop-down list) when the text would
 * land inside them; at their edge the text is typed beside the control instead. `undefined`
 * means the position is not in a placeholder and ordinary typing applies.
 */
export function typeOverPlaceholder(
	state: EditorState,
	from: number,
	text: string,
	to = from,
): Transaction | 'blocked' | undefined {
	const span = inputControlAt(controlSpans(state.doc), from);
	if (span?.level === 'cell') return typeInCell(state, span, text, to);
	if (span?.level !== 'inline') return undefined;
	const { control } = span;
	if (control.kind === 'checkbox' || control.kind === 'dropDownList') {
		const inside = (from > span.from && from < span.to) || (from === span.from && to === span.to);
		if (inside || (from === span.from && to > from)) return 'blocked';
		// Beside the control: the new text must not become part of it.
		const type = state.schema.marks.contentControl!;
		return state.tr.insertText(text, from).removeMark(from, from + text.length, type);
	}
	if (!control.showingPlaceholder) return undefined;
	if (!editable(state, span)) return 'blocked';
	const tr = replaceControlContent(state, span, afterEdit({ control: withValue(control), text }));
	const end = tr.mapping.map(span.to);
	return tr.setSelection(TextSelection.near(tr.doc.resolve(end)));
}

/**
 * Typing in a cell control: a list or check box takes no text, a placeholder is replaced whole. A
 * selection that runs on past the cell (`to`) is typed over as ordinary text, as it is deleted.
 */
function typeInCell(
	state: EditorState,
	span: ControlSpan,
	text: string,
	to: number,
): Transaction | 'blocked' | undefined {
	const { control } = span;
	if (control.kind === 'checkbox' || control.kind === 'dropDownList') return 'blocked';
	if (to > span.to - 1 || !control.showingPlaceholder || !replaceable(state, span))
		return undefined;
	if (!editable(state, span)) return 'blocked';
	const tr = replaceControlContent(state, span, afterEdit({ control: withValue(control), text }));
	return tr.setSelection(
		TextSelection.create(tr.doc, cellText(tr.doc, span.from).from + text.length),
	);
}

/** `fragment` without the Placeholder Text style, which pasted plain text takes where it lands. */
function withoutPlaceholderStyle(fragment: Fragment): Fragment {
	const nodes: ProseMirrorNode[] = [];
	fragment.forEach((node) =>
		nodes.push(
			node.isText
				? node.mark(node.marks.filter((mark) => !isPlaceholderStyle(mark)))
				: node.copy(withoutPlaceholderStyle(node.content)),
		),
	);
	return Fragment.fromArray(nodes);
}

/**
 * A paste in a cell control is input, as typing is: a list or check box cell takes none, and a
 * placeholder is replaced whole, the control no longer showing it. `undefined` means the paste is
 * not in a cell control, runs on past it, or goes in as an ordinary paste.
 */
export function pasteInCell(state: EditorState, slice: Slice): Transaction | 'blocked' | undefined {
	const { from, to } = state.selection;
	const span = inputControlAt(controlSpans(state.doc), from);
	if (span?.level !== 'cell' || to > span.to - 1) return undefined;
	const { control } = span;
	if (control.kind === 'checkbox' || control.kind === 'dropDownList') return 'blocked';
	if (!control.showingPlaceholder || !replaceable(state, span)) return undefined;
	if (!editable(state, span)) return 'blocked';
	// The placeholder goes and an empty paragraph takes the pasted content, as typed text would.
	const tr = replaceControlContent(
		state,
		span,
		afterEdit({ control: withValue(control), text: '' }),
	);
	const at = TextSelection.create(tr.doc, cellText(tr.doc, span.from).from);
	return tr
		.setSelection(at)
		.replaceSelection(
			new Slice(withoutPlaceholderStyle(slice.content), slice.openStart, slice.openEnd),
		);
}

/** Deleting everything a text control holds brings its placeholder back instead of removing it. */
export function deleteToPlaceholder(
	state: EditorState,
	direction: 'backward' | 'forward',
): Transaction | undefined {
	const { selection } = state;
	const span = inputControlAt(controlSpans(state.doc), selection.from);
	if (!span || span.level === 'block' || span.control.showingPlaceholder) return undefined;
	if (!editable(state, span) || span.control.kind === 'checkbox') return undefined;
	if (!replaceable(state, span)) return undefined;
	// Deleting a temporary control's text removes the control instead (content-control-temporary).
	if (removedByEdit(span.control)) return undefined;
	const [from, to] = selection.empty
		? direction === 'backward'
			? [selection.from - 1, selection.from]
			: [selection.from, selection.from + 1]
		: [selection.from, selection.to];
	// The deletion takes all the control holds when nothing inline stays before or after it in the
	// control's text (`cellText` for a cell, whose other paragraphs can be empty). A selection that
	// covers that text and goes on past the control is deleted as ordinary text, all of it.
	const cell = span.level === 'cell';
	const inner = cell ? cellText(state.doc, span.from) : span;
	const outer = cell ? { from: span.from + 1, to: span.to - 1 } : span;
	if (from < outer.from || to > outer.to) return undefined;
	// A table in a cell stays when its placeholder comes back: deleting in one is ordinary deleting.
	if (cell && touchesTable(state.doc, span.from, from, to)) return undefined;
	if (!holdsNothing(state.doc, inner.from, from) || !holdsNothing(state.doc, to, inner.to))
		return undefined;
	return replaceControlContent(state, span, placeholderChange(span.control));
}

/**
 * A new control id that no control in `doc` has, rows and tables included: Word uses signed 32-bit
 * integers.
 */
export const newControlId = (doc: ProseMirrorNode): string =>
	freshControlId(new Set(everyControlSpan(doc).map((span) => span.control.id)));

export interface NewControlOptions {
	alias?: string;
	tag?: string;
	items?: { display: string; value: string }[];
	dateFormat?: string;
}

/** Insert > Content Control: wraps the selected text, or inserts a placeholder (or a checkbox). */
export function insertContentControl(
	state: EditorState,
	kind: ContentControlKind,
	options: NewControlOptions = {},
): Transaction | undefined {
	const type = state.schema.marks.contentControl;
	if (!type) return undefined;
	const { selection } = state;
	const base: ContentControl = {
		id: newControlId(state.doc),
		level: 'inline',
		kind,
		...(options.alias ? { alias: options.alias } : {}),
		...(options.tag ? { tag: options.tag } : {}),
		...(options.items ? { items: options.items } : {}),
		...(kind === 'date'
			? { dateFormat: options.dateFormat ?? 'M/d/yyyy', dateLocale: 'en-US' }
			: {}),
	};
	const wrapsText = !selection.empty && (kind === 'richText' || kind === 'text');
	if (wrapsText) {
		return state.tr
			.addMark(selection.from, selection.to, type.create({ data: JSON.stringify(base) }))
			.setMeta(CONTROL_EDIT_META, true);
	}
	const change =
		kind === 'checkbox'
			? checkboxState({ ...base, checked: false }, false)
			: placeholderChange(base);
	const marks = [type.create({ data: JSON.stringify(change.control) })];
	if (kind !== 'checkbox' && state.schema.marks.characterStyle)
		marks.push(state.schema.marks.characterStyle.create({ id: PLACEHOLDER_STYLE }));
	if (change.font && state.schema.marks.font)
		marks.push(state.schema.marks.font.create({ family: change.font, size: null, color: null }));
	const start = state.selection.from;
	const tr = state.tr.replaceSelectionWith(state.schema.text(change.text, marks), false);
	// Word selects a new control's placeholder, so the first typed character replaces it.
	if (kind !== 'checkbox')
		tr.setSelection(TextSelection.create(tr.doc, start, start + change.text.length));
	return tr.setMeta(CONTROL_EDIT_META, true);
}
