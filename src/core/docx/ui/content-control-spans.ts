// Where content controls sit in the editor document: inline controls are runs of text sharing a
// `contentControl` mark, block controls are consecutive paragraphs sharing a node attribute.
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type {
	ContentControl,
	ContentControlGlyph,
	ContentControlListItem,
} from '../content-control-model';

/** A control and the document range it covers. */
export interface ControlSpan {
	control: ContentControl;
	level: 'inline' | 'block';
	/** Inline: the first text position. Block: the position before the first paragraph. */
	from: number;
	/** Inline: after the last character. Block: after the last paragraph. */
	to: number;
}

/** The fields of a control that hold text, and those that hold a flag. */
const TEXT_FIELDS = [
	'alias',
	'tag',
	'lock',
	'dateFormat',
	'dateLocale',
	'fullDate',
	'placeholder',
	'placeholderText',
	'sourceXml',
	'endSourceXml',
] as const;
const FLAG_FIELDS = [
	'showingPlaceholder',
	'multiLine',
	'checked',
	'dataBound',
	'temporary',
] as const;

const isGlyph = (value: unknown): value is ContentControlGlyph => {
	const { code, font } = (value ?? {}) as { code?: unknown; font?: unknown };
	return typeof code === 'string' && (font === undefined || typeof font === 'string');
};
const isItem = (value: unknown): value is ContentControlListItem => {
	const { display, value: entry } = (value ?? {}) as { display?: unknown; value?: unknown };
	return typeof display === 'string' && typeof entry === 'string';
};
/** A list entry's text: text, or a number written as text. */
const entryText = (value: unknown): string | undefined =>
	typeof value === 'string'
		? value
		: typeof value === 'number' && Number.isFinite(value)
			? String(value)
			: undefined;

/**
 * `control` with only values of the kinds its fields hold: text, flags, glyphs with a text code,
 * list entries with a text display and value. A page can put any JSON in `data-sdt`, and an object
 * where the editor draws or the writer saves text can make the paste fail to draw or the document
 * fail to save. Itself when every value is of its kind.
 */
function wellFormed(control: ContentControl): ContentControl {
	const fields = control as unknown as Record<string, unknown>;
	const dropped = [
		...TEXT_FIELDS.filter(
			(field) => fields[field] !== undefined && typeof fields[field] !== 'string',
		),
		...FLAG_FIELDS.filter(
			(field) => fields[field] !== undefined && typeof fields[field] !== 'boolean',
		),
		...(['checkedGlyph', 'uncheckedGlyph'] as const).filter(
			(field) => fields[field] !== undefined && !isGlyph(fields[field]),
		),
	];
	const items = fields.items;
	const listed = items === undefined || (Array.isArray(items) && items.every(isItem));
	if (!dropped.length && listed) return control;
	const kept: Record<string, unknown> = { ...fields };
	for (const field of dropped) delete kept[field];
	if (!listed) {
		if (!Array.isArray(items)) delete kept.items;
		else
			kept.items = items.flatMap((item: unknown) => {
				const { display, value } = (item ?? {}) as { display?: unknown; value?: unknown };
				const [shown, entry] = [entryText(display), entryText(value)];
				const text = shown ?? entry;
				return text === undefined ? [] : [{ display: text, value: entry ?? text }];
			});
	}
	return kept as unknown as ContentControl;
}

/**
 * Whether `value` nests objects and arrays no deeper than a control does, the fields of its list
 * entries included, checked level by level: deeper nesting can make turning it back into JSON
 * overflow the stack, in a paste or in every later save.
 */
function nestsLikeControl(value: unknown): boolean {
	const isNode = (item: unknown): item is object => Boolean(item) && typeof item === 'object';
	let level = [value];
	for (let depth = 0; level.length; depth++) {
		if (depth > 3) return false;
		level = level.flatMap((item) => (isNode(item) ? Object.values(item).filter(isNode) : []));
	}
	return true;
}

/**
 * The inline control a `contentControl` mark holds. HTML pasted from another page can put any JSON
 * in its `data-sdt`: only an object with an id and a kind, nested no deeper than a control's
 * fields, is a control, and it holds only values of the kinds its fields hold.
 */
export function controlOfMark(mark: Mark | undefined): ContentControl | undefined {
	if (!mark?.attrs.data) return undefined;
	let value: unknown;
	try {
		value = JSON.parse(String(mark.attrs.data));
	} catch {
		return undefined;
	}
	if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
	if (!nestsLikeControl(value)) return undefined;
	const { id, kind } = value as { id?: unknown; kind?: unknown };
	if (typeof id !== 'string' || typeof kind !== 'string') return undefined;
	return wellFormed(value as ContentControl);
}

const markOf = (node: ProseMirrorNode): Mark | undefined =>
	node.marks.find((mark) => mark.type.name === 'contentControl');

/** Every inline and block control in document order. */
export function controlSpans(doc: ProseMirrorNode): ControlSpan[] {
	const spans: ControlSpan[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name === 'paragraph' && node.attrs.contentControl) {
			const control = node.attrs.contentControl as ContentControl;
			const previous = spans.at(-1);
			if (previous?.level === 'block' && previous.control.id === control.id && previous.to === pos)
				previous.to = pos + node.nodeSize;
			else spans.push({ control, level: 'block', from: pos, to: pos + node.nodeSize });
		}
		if (!node.isInline) return true;
		const control = controlOfMark(markOf(node));
		if (!control) return false;
		const previous = spans.at(-1);
		if (previous?.level === 'inline' && previous.control.id === control.id && previous.to === pos) {
			previous.to = pos + node.nodeSize;
			previous.control = control;
		} else spans.push({ control, level: 'inline', from: pos, to: pos + node.nodeSize });
		return false;
	});
	return spans;
}

/** The innermost control around `pos` (an inline control wins over the block one around it). */
export function controlAt(
	spans: readonly ControlSpan[],
	pos: number,
	edges = false,
): ControlSpan | undefined {
	const inside = spans.filter((span) => pos > span.from && pos < span.to);
	const found = inside.find((span) => span.level === 'inline') ?? inside[0];
	if (found || !edges) return found;
	return spans.find((span) => span.level === 'inline' && (pos === span.from || pos === span.to));
}

export const controlAtSelection = (state: EditorState, edges = true): ControlSpan | undefined =>
	controlAt(controlSpans(state.doc), state.selection.from, edges);
