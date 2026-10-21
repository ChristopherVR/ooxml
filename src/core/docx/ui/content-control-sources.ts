// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// A control's verbatim properties (`sourceXml`) come with it when it is copied: the editor keeps
// an inline control in its span's `data-sdt` attribute, so HTML from another page can bring any.
// Pasted ones stay only when this editor is where they come from: the document holds them (a copy
// of one of its own controls), or a copy, cut or drag in this editor took them (a move). A control
// from elsewhere is written from its modeled fields, and keeps only values of the kinds they hold.
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import { controlOfMark, controlSpans } from './content-control-spans';

/** The verbatim properties of the inline controls in `slice`. */
export function controlSourcesOf(slice: Slice): string[] {
	const sources: string[] = [];
	slice.content.descendants((node) => {
		for (const mark of node.marks) {
			const source =
				mark.type.name === 'contentControl' ? controlOfMark(mark)?.sourceXml : undefined;
			if (source) sources.push(source);
		}
	});
	return sources;
}

/**
 * `slice` with the verbatim properties taken off its controls that neither `doc` holds nor a copy,
 * cut or drag in the same editor took (`copied`), and its controls holding only values of the
 * kinds their fields hold (`controlOfMark`): a page can put any JSON in `data-sdt`, and an object
 * where text is drawn or saved would make the paste fail to draw or the document fail to save. A
 * mark that holds no control is taken off; its text stays.
 */
export function pastedControlSources(
	slice: Slice,
	doc: ProseMirrorNode,
	copied: ReadonlySet<string> = new Set(),
): Slice {
	let held: Set<string> | undefined;
	const mapMark = (mark: Mark): Mark | undefined => {
		if (mark.type.name !== 'contentControl') return mark;
		let control = controlOfMark(mark);
		if (!control) return undefined;
		if (control.sourceXml && !copied.has(control.sourceXml)) {
			held ??= new Set(controlSpans(doc).flatMap((span) => span.control.sourceXml ?? []));
			if (!held.has(control.sourceXml)) {
				const { sourceXml: _untrusted, ...modeled } = control;
				control = modeled;
			}
		}
		let data: string;
		try {
			data = JSON.stringify(control);
		} catch {
			return undefined;
		}
		return data === mark.attrs.data ? mark : mark.type.create({ ...mark.attrs, data });
	};
	const mapNode = (node: ProseMirrorNode): ProseMirrorNode => {
		const marks = node.marks.flatMap((mark) => mapMark(mark) ?? []);
		const marked =
			marks.length !== node.marks.length || marks.some((mark, index) => mark !== node.marks[index])
				? node.mark(marks)
				: node;
		const content = mapFragment(node.content);
		return content === node.content ? marked : marked.copy(content);
	};
	const mapFragment = (fragment: Fragment): Fragment => {
		const nodes: ProseMirrorNode[] = [];
		let changed = false;
		fragment.forEach((child) => {
			const next = mapNode(child);
			changed ||= next !== child;
			nodes.push(next);
		});
		return changed ? Fragment.fromArray(nodes) : fragment;
	};
	const content = mapFragment(slice.content);
	return content === slice.content ? slice : new Slice(content, slice.openStart, slice.openEnd);
}
