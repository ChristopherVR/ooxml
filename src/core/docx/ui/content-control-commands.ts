// Editing commands for content controls: toggling, picking, typing over a placeholder, inserting.
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
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
import type { ContentControl, ContentControlKind } from '../content-control-model';
import { freshControlId } from './content-control-paste';
import { controlAt, controlSpans, type ControlSpan } from './content-control-spans';

/** Transactions carrying this meta are the control's own edits and bypass the lock guard. */
export const CONTROL_EDIT_META = 'dve-content-control';

const PLACEHOLDER_STYLE = 'PlaceholderText';
const isPlaceholderStyle = (mark: Mark) =>
	mark.type.name === 'characterStyle' && mark.attrs.id === PLACEHOLDER_STYLE;

/** The first text of a control: its formatting carries over to the content that replaces it. */
function firstText(state: EditorState, span: ControlSpan): ProseMirrorNode | null {
	if (span.level === 'inline') return state.doc.nodeAt(span.from);
	let found: ProseMirrorNode | null = null;
	state.doc.nodesBetween(span.from, span.to, (node) => {
		if (!found && node.isText) found = node;
		return !found;
	});
	return found;
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

/**
 * Replaces a control's content (and its state) as one edit. An inline control's text is replaced;
 * a cell control's cell keeps one paragraph, its first, holding the new text.
 */
export function replaceControlContent(
	state: EditorState,
	span: ControlSpan,
	change: ControlChange,
): Transaction {
	const marks = contentMarks(state, span, change);
	const text = change.text ? state.schema.text(change.text, marks) : null;
	if (span.level !== 'cell')
		return state.tr.replaceWith(span.from, span.to, text ?? []).setMeta(CONTROL_EDIT_META, true);
	const cell = state.doc.nodeAt(span.from)!;
	const first = cell.firstChild?.type.name === 'paragraph' ? cell.firstChild : undefined;
	const paragraph = (first?.type ?? state.schema.nodes.paragraph!).create(first?.attrs, text);
	const tr = state.tr.replaceWith(span.from + 1, span.to - 1, paragraph);
	tr.setNodeMarkup(span.from, undefined, {
		...cell.attrs,
		contentControl: change.removed ? null : change.control,
	});
	// A caret in the cell stays in it: mapped past the replaced content it would land in the next
	// cell. It goes to the start of a placeholder (typing replaces it), else after the text.
	const { from, to } = state.selection;
	if (from > span.from && to < span.to) {
		const shown = change.control.showingPlaceholder && !change.removed;
		const start = span.from + 2;
		tr.setSelection(TextSelection.create(tr.doc, shown ? start : start + (text?.nodeSize ?? 0)));
	}
	return tr.setMeta(CONTROL_EDIT_META, true);
}

/**
 * Whether the content of `span` may be replaced as a whole: not when it holds a control that may
 * not be deleted (`sdtLocked`), since the replacement carries `CONTROL_EDIT_META` past the lock guard.
 */
function replaceable(state: EditorState, span: ControlSpan): boolean {
	return !controlSpans(state.doc).some(
		(inner) =>
			inner.from >= span.from &&
			inner.to <= span.to &&
			(inner.level !== span.level || inner.from !== span.from) &&
			!canDeleteControl(inner.control),
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
		!canEditControlContent(span.control) ||
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
	if (!span || !canEditControlContent(span.control) || !replaceable(state, span)) return undefined;
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
	if (
		!span ||
		span.control.kind !== 'date' ||
		!canEditControlContent(span.control) ||
		!replaceable(state, span)
	)
		return undefined;
	return replaceControlContent(state, span, setControlDate(span.control, date));
}

const placeholderChange = (control: ContentControl): ControlChange => ({
	control: { ...control, showingPlaceholder: true },
	text: placeholderTextOf(control),
});

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
	const span = controlAt(controlSpans(state.doc), from, true);
	if (span?.level === 'cell') return typeInCell(state, span, text);
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
	if (!canEditControlContent(control)) return 'blocked';
	const tr = replaceControlContent(state, span, afterEdit({ control: withValue(control), text }));
	const end = tr.mapping.map(span.to);
	return tr.setSelection(TextSelection.near(tr.doc.resolve(end)));
}

/** Typing in a cell control: a list or check box takes no text, a placeholder is replaced whole. */
function typeInCell(
	state: EditorState,
	span: ControlSpan,
	text: string,
): Transaction | 'blocked' | undefined {
	const { control } = span;
	if (control.kind === 'checkbox' || control.kind === 'dropDownList') return 'blocked';
	if (!control.showingPlaceholder || !replaceable(state, span)) return undefined;
	if (!canEditControlContent(control)) return 'blocked';
	const tr = replaceControlContent(state, span, afterEdit({ control: withValue(control), text }));
	return tr.setSelection(TextSelection.create(tr.doc, span.from + 2 + text.length));
}

/** Deleting everything a text control holds brings its placeholder back instead of removing it. */
export function deleteToPlaceholder(
	state: EditorState,
	direction: 'backward' | 'forward',
): Transaction | undefined {
	const { selection } = state;
	const span = controlAt(controlSpans(state.doc), selection.from, true);
	if (!span || span.level === 'block' || span.control.showingPlaceholder) return undefined;
	if (!canEditControlContent(span.control) || span.control.kind === 'checkbox') return undefined;
	if (!replaceable(state, span)) return undefined;
	// Deleting a temporary control's text removes the control instead (content-control-temporary).
	if (removedByEdit(span.control)) return undefined;
	const [from, to] = selection.empty
		? direction === 'backward'
			? [selection.from - 1, selection.from]
			: [selection.from, selection.from + 1]
		: [selection.from, selection.to];
	// A cell's text runs from inside its first paragraph to the end of its last. A selection that
	// covers it and goes on past the control is deleted as ordinary text, all of it.
	const cell = span.level === 'cell';
	const inner = cell ? { from: span.from + 2, to: span.to - 2 } : span;
	const outer = cell ? { from: span.from + 1, to: span.to - 1 } : span;
	if (from > inner.from || to < inner.to || from < outer.from || to > outer.to) return undefined;
	return replaceControlContent(state, span, placeholderChange(span.control));
}

/** A new control id: Word uses signed 32-bit integers. */
export const newControlId = (doc: ProseMirrorNode): string =>
	freshControlId(new Set(controlSpans(doc).map((span) => span.control.id)));

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
