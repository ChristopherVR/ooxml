// Editing commands for content controls: toggling, picking, typing over a placeholder, inserting.
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import {
	canEditControlContent,
	checkboxState,
	defaultPlaceholder,
	selectListItem,
	setControlDate,
	toggleCheckbox,
	withValue,
	type ControlChange,
} from '../content-control-edit';
import type { ContentControl, ContentControlKind } from '../content-control-model';
import { controlAt, controlSpans, type ControlSpan } from './content-control-spans';

/** Transactions carrying this meta are the control's own edits and bypass the lock guard. */
export const CONTROL_EDIT_META = 'dve-content-control';

const PLACEHOLDER_STYLE = 'PlaceholderText';
const isPlaceholderStyle = (mark: Mark) =>
	mark.type.name === 'characterStyle' && mark.attrs.id === PLACEHOLDER_STYLE;

function contentMarks(state: EditorState, span: ControlSpan, change: ControlChange): Mark[] {
	const schema = state.schema;
	const first = state.doc.nodeAt(span.from);
	const kept = (first?.marks ?? []).filter(
		(mark) => mark.type.name !== 'contentControl' && !isPlaceholderStyle(mark),
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
	marks.push(schema.marks.contentControl!.create({ data: JSON.stringify(change.control) }));
	return marks;
}

/** Replaces an inline control's content (and its state) as one edit. */
export function replaceControlContent(
	state: EditorState,
	span: ControlSpan,
	change: ControlChange,
): Transaction {
	const tr = state.tr.replaceWith(
		span.from,
		span.to,
		state.schema.text(change.text, contentMarks(state, span, change)),
	);
	return tr.setMeta(CONTROL_EDIT_META, true);
}

/** Stores new properties on every run or paragraph of the control. */
export function updateControlProperties(
	state: EditorState,
	span: ControlSpan,
	control: ContentControl,
): Transaction {
	const tr = state.tr.setMeta(CONTROL_EDIT_META, true);
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

const spanById = (state: EditorState, id: string): ControlSpan | undefined =>
	controlSpans(state.doc).find((span) => span.level === 'inline' && span.control.id === id);

export function toggleCheckboxControl(state: EditorState, id: string): Transaction | undefined {
	const span = spanById(state, id);
	if (!span || span.control.kind !== 'checkbox' || !canEditControlContent(span.control))
		return undefined;
	return replaceControlContent(state, span, toggleCheckbox(span.control));
}

export function chooseListItem(
	state: EditorState,
	id: string,
	index: number,
): Transaction | undefined {
	const span = spanById(state, id);
	const change = span && canEditControlContent(span.control) && selectListItem(span.control, index);
	return span && change ? replaceControlContent(state, span, change) : undefined;
}

export function chooseDate(state: EditorState, id: string, date: Date): Transaction | undefined {
	const span = spanById(state, id);
	if (!span || span.control.kind !== 'date' || !canEditControlContent(span.control))
		return undefined;
	return replaceControlContent(state, span, setControlDate(span.control, date));
}

const placeholderChange = (control: ContentControl): ControlChange => ({
	control: { ...control, showingPlaceholder: true },
	text: defaultPlaceholder(control),
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
	const tr = replaceControlContent(state, span, { control: withValue(control), text });
	const end = tr.mapping.map(span.to);
	return tr.setSelection(TextSelection.near(tr.doc.resolve(end)));
}

/** Deleting everything a text control holds brings its placeholder back instead of removing it. */
export function deleteToPlaceholder(
	state: EditorState,
	direction: 'backward' | 'forward',
): Transaction | undefined {
	const { selection } = state;
	const span = controlAt(controlSpans(state.doc), selection.from, true);
	if (span?.level !== 'inline' || span.control.showingPlaceholder) return undefined;
	if (!canEditControlContent(span.control) || span.control.kind === 'checkbox') return undefined;
	const [from, to] = selection.empty
		? direction === 'backward'
			? [selection.from - 1, selection.from]
			: [selection.from, selection.from + 1]
		: [selection.from, selection.to];
	if (from > span.from || to < span.to) return undefined;
	return replaceControlContent(state, span, placeholderChange(span.control));
}

/** A new control id: Word uses signed 32-bit integers. */
export const newControlId = (doc: ProseMirrorNode): string => {
	const used = new Set(controlSpans(doc).map((span) => span.control.id));
	let id: string;
	do id = String(Math.floor(Math.random() * 2 ** 31) + 1);
	while (used.has(id));
	return id;
};

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
