import type { NodeSpec } from 'prosemirror-model';
import { inlineRunFormattingAttrs, inlineRunPropertiesDomAttrs } from './inline-run-properties';

/** A line break carries its own run properties through element-based collaboration mappings. */
export const hardBreakNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: false,
	attrs: { ...inlineRunFormattingAttrs, format: { default: null } },
	leafText: () => '\n',
	parseDOM: [
		{
			tag: 'br',
			getAttrs: (el) => ({ format: (el as HTMLElement).dataset.runProperties ?? null }),
		},
	],
	toDOM: (node) => ['br', inlineRunPropertiesDomAttrs(node)],
};

/** A visible, editable page or column break marker (`w:br` type page/column). */
export const pageBreakNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: true,
	attrs: { ...inlineRunFormattingAttrs, kind: { default: 'page' }, format: { default: null } },
	leafText: () => '',
	parseDOM: [
		{
			tag: 'span[data-docx-break]',
			getAttrs: (el) => ({
				kind: (el as HTMLElement).dataset.docxBreak === 'column' ? 'column' : 'page',
				format: (el as HTMLElement).dataset.runProperties ?? null,
			}),
		},
	],
	toDOM: (node) => [
		'span',
		{
			class: 'dve-break-marker',
			'data-docx-break': node.attrs.kind,
			contenteditable: 'false',
			...inlineRunPropertiesDomAttrs(node),
		},
		node.attrs.kind === 'column' ? 'Column Break' : 'Page Break',
	],
};

/** A read-only footnote/endnote reference mark; its number is derived from document order. */
export const noteReferenceNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: false,
	attrs: {
		...inlineRunFormattingAttrs,
		kind: { default: 'footnote' },
		id: { default: '' },
		number: { default: 1 },
		/** The number as shown, in the document's footnote/endnote number format. */
		label: { default: null },
		/** The reference run's own formatting (e.g. superscript, FootnoteReference style) as JSON.
		 *  Kept as an attribute, not marks, so text typed beside a reference doesn't inherit it. */
		format: { default: null },
	},
	leafText: (node) => String(node.attrs.label ?? node.attrs.number),
	parseDOM: [
		{
			tag: 'sup[data-docx-note-kind]',
			priority: 60,
			getAttrs: (el) => ({
				kind: (el as HTMLElement).dataset.docxNoteKind === 'endnote' ? 'endnote' : 'footnote',
				format: (el as HTMLElement).dataset.runProperties ?? null,
				id: (el as HTMLElement).dataset.docxNoteId || '',
				number: Number((el as HTMLElement).textContent) || 1,
			}),
		},
	],
	toDOM: (node) => [
		'sup',
		{
			class: 'dve-note-reference',
			'data-docx-note-kind': node.attrs.kind,
			'data-docx-note-id': node.attrs.id,
			...inlineRunPropertiesDomAttrs(node),
			contenteditable: 'false',
		},
		String(node.attrs.label ?? node.attrs.number),
	],
};

/**
 * A complex field's structural run: a `w:fldChar` marker (begin/separate/end) or its instruction
 * text (`code`). Invisible while editing; its run formatting rides in `format` like note references.
 */
export const fieldMarkerNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: false,
	attrs: {
		...inlineRunFormattingAttrs,
		kind: { default: 'begin' },
		code: { default: null },
		format: { default: null },
	},
	leafText: () => '',
	parseDOM: [
		{
			tag: 'span[data-field-marker]',
			getAttrs: (el) => ({
				kind: (el as HTMLElement).dataset.fieldMarker || 'begin',
				format: (el as HTMLElement).dataset.runProperties ?? null,
				code: (el as HTMLElement).dataset.fieldCode ?? null,
			}),
		},
	],
	toDOM: (node) => [
		'span',
		{
			class: 'dve-field-marker',
			'data-field-marker': node.attrs.kind,
			...inlineRunPropertiesDomAttrs(node),
			...(node.attrs.code != null ? { 'data-field-code': node.attrs.code } : {}),
			contenteditable: 'false',
		},
	],
};
