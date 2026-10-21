// A pasted content control gets a new id when the document already has a control with its id.
// Ids are unique in a document (ECMA-376 17.5.2.18: Word gives a repeated one a new id when it
// opens the file), and a copy pasted beside its source must be operated and saved as itself.
// Text pasted inside a control joins it, as typed text does, so that control keeps its id. Block
// controls keep their ids: the writer joins a new block to the control of the block before it by
// id, as it does for a paragraph typed at the end of one.
// A pasted control keeps its verbatim properties only when this editor is where they come from, the
// rule content-control-sources.ts applies to an inline control's `w:sdtPr`. HTML from another page
// brings more: a cell's control in `data-sdt`, and in its slice context (`data-pm-slice`) any
// attributes for the nodes around what it holds, which no parse rule reads. So the control of a
// pasted paragraph, table, row or cell is taken only when it is one, at its level, and its verbatim
// properties, as every control's end properties (`w:sdtEndPr`), follow the same rule.
import {
	Fragment,
	Slice,
	type Attrs,
	type Mark,
	type Node as ProseMirrorNode,
} from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { freshControlId, type ContentControl } from '../content-control-model';
import { rowControlOf } from './content-control-rows';
import { controlFromValue, controlOfMark, everyControlSpan } from './content-control-spans';

type Verbatim = 'sourceXml' | 'endSourceXml';
const VERBATIM: readonly Verbatim[] = ['sourceXml', 'endSourceXml'];
/** What is checked here of an inline control: `pastedControlSources` checks its `sourceXml`. */
const INLINE: readonly Verbatim[] = ['endSourceXml'];

const sourcesOf = (
	control: ContentControl | undefined,
	fields: readonly Verbatim[] = VERBATIM,
): string[] =>
	fields.flatMap((field) => {
		const source: unknown = control?.[field];
		return typeof source === 'string' ? [source] : [];
	});

/** What `node` holds as its control: a paragraph's or table's block control, a row's or a cell's. */
export function nodeControlOf(
	node: ProseMirrorNode,
): { value: unknown; level: ContentControl['level'] } | undefined {
	switch (node.type.name) {
		case 'paragraph':
		case 'table':
			return { value: node.attrs.contentControl, level: 'block' };
		case 'tableCell':
			return { value: node.attrs.contentControl, level: 'cell' };
		case 'tableRow':
			return {
				value: (node.attrs.properties as { contentControl?: unknown } | null)?.contentControl,
				level: 'row',
			};
		default:
			return undefined;
	}
}

/** The attributes of `node` with `control` as the control it holds, or with none. */
function attrsWithControl(node: ProseMirrorNode, control: ContentControl | undefined): Attrs {
	if (node.type.name !== 'tableRow') return { ...node.attrs, contentControl: control ?? null };
	const { contentControl: _held, ...properties } = (node.attrs.properties ?? {}) as Record<
		string,
		unknown
	>;
	return {
		...node.attrs,
		properties: control
			? { ...properties, contentControl: control }
			: Object.keys(properties).length
				? properties
				: null,
	};
}

/**
 * The verbatim properties in `slice` other than those `controlSourcesOf` takes, for a copy, cut or
 * drag in the editor to record: every block, row and cell control's, and every control's end
 * properties.
 */
export function otherControlSourcesOf(slice: Slice): string[] {
	const sources: string[] = [];
	slice.content.descendants((node) => {
		const holds = nodeControlOf(node);
		if (holds) sources.push(...sourcesOf(controlFromValue(holds.value, holds.level)));
		for (const mark of node.marks)
			if (mark.type.name === 'contentControl')
				sources.push(...sourcesOf(controlOfMark(mark), INLINE));
	});
	return sources;
}

/**
 * `slice` as it is pasted in `view`: its inline, cell and row controls under ids the document
 * lacks, its nodes holding only what is a control, and without the verbatim properties checked
 * here that neither the document holds nor a copy, cut or drag in the editor took (`copied`). Plain
 * text (`plain`) takes the marks of where it lands, and a drag that moves content takes its
 * controls along: neither holds a copy of a control.
 */
export function pastedControlSlice(
	slice: Slice,
	view: EditorView | undefined,
	plain = false,
	copied: ReadonlySet<string> = new Set(),
): Slice {
	if (!view || plain || view.dragging?.move) return slice;
	const { doc, selection, storedMarks } = view.state;
	const spans = everyControlSpan(doc);
	const inDocument = new Set(spans.map((span) => span.control.id));
	const held = new Set(spans.flatMap((span) => sourcesOf(span.control)));
	/** `control` without those of its `fields` that come from elsewhere: the model is written. */
	const withOwnSources = (control: ContentControl, fields: readonly Verbatim[]): ContentControl => {
		const foreign = fields.filter((field) => {
			const source: unknown = control[field];
			return (
				source !== undefined &&
				!(typeof source === 'string' && (held.has(source) || copied.has(source)))
			);
		});
		if (!foreign.length) return control;
		const kept = { ...control };
		for (const field of foreign) delete kept[field];
		return kept;
	};
	// The inline controls pasted text joins: those whose mark text typed there would take. A
	// drag's selection is where it came from, not where it lands, so a dropped copy joins none.
	const landing = view.dragging
		? []
		: selection.empty
			? (storedMarks ?? selection.$from.marks())
			: (selection.$from.marksAcross(selection.$to) ?? []);
	const joined = new Set(
		landing.flatMap((mark) => {
			const id = mark.type.name === 'contentControl' ? controlOfMark(mark)?.id : undefined;
			return id ? [id] : [];
		}),
	);
	// A new id is new to the document and to the pasted content both.
	const taken = new Set(inDocument);
	slice.content.descendants((node) => {
		for (const mark of node.marks) {
			const id = mark.type.name === 'contentControl' ? controlOfMark(mark)?.id : undefined;
			if (id) taken.add(id);
		}
		const id = (node.attrs.contentControl as ContentControl | null | undefined)?.id;
		const row = rowControlOf(node)?.id;
		for (const found of [id, row]) if (found) taken.add(found);
	});
	const renamed = new Map<string, string>();
	const withNewId = (control: ContentControl): ContentControl => {
		if (!inDocument.has(control.id)) return control;
		let id = renamed.get(control.id);
		if (!id) {
			id = freshControlId(taken);
			taken.add(id);
			renamed.set(control.id, id);
		}
		return { ...control, id };
	};
	/** The control a pasted node holds, as it is taken; block controls keep their ids. */
	const pastedControl = (
		value: unknown,
		level: ContentControl['level'],
	): ContentControl | undefined => {
		const control = controlFromValue(value, level);
		if (!control) return undefined;
		const kept = withOwnSources(control, VERBATIM);
		return level === 'block' ? kept : withNewId(kept);
	};
	const markOf = (mark: Mark): Mark => {
		const pasted = mark.type.name === 'contentControl' ? controlOfMark(mark) : undefined;
		if (!pasted) return mark;
		const kept = withOwnSources(pasted, INLINE);
		const control = joined.has(pasted.id) ? kept : withNewId(kept);
		return control === pasted
			? mark
			: mark.type.create({ ...mark.attrs, data: JSON.stringify(control) });
	};
	const nodeOf = (node: ProseMirrorNode): ProseMirrorNode => {
		if (node.isInline) {
			const marks = node.marks.map(markOf);
			return marks.some((mark, index) => mark !== node.marks[index]) ? node.mark(marks) : node;
		}
		const content = fragmentOf(node.content);
		const holds = nodeControlOf(node);
		const pasted = holds?.value == null ? undefined : pastedControl(holds.value, holds.level);
		const changed = holds?.value != null && pasted !== holds.value;
		if (!changed && content === node.content) return node;
		return node.type.create(
			changed ? attrsWithControl(node, pasted) : node.attrs,
			content,
			node.marks,
		);
	};
	const fragmentOf = (fragment: Fragment): Fragment => {
		const nodes: ProseMirrorNode[] = [];
		let changed = false;
		fragment.forEach((node) => {
			const next = nodeOf(node);
			changed ||= next !== node;
			nodes.push(next);
		});
		return changed ? Fragment.fromArray(nodes) : fragment;
	};
	const content = fragmentOf(slice.content);
	return content === slice.content ? slice : new Slice(content, slice.openStart, slice.openEnd);
}
