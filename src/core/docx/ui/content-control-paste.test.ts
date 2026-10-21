import { describe, expect, it, vi } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { ContentControl } from '../model';
import { markSpecs } from './schema-marks';
import { contentControlPlugin } from './content-control-guard';
import { freshControlId } from '../content-control-model';
import { controlSpans } from './content-control-spans';

const schema = new Schema({
	nodes: {
		doc: { content: 'block+' },
		paragraph: {
			group: 'block',
			content: 'inline*',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['p', 0],
		},
		table: { group: 'block', content: 'tableRow+', toDOM: () => ['table', ['tbody', 0]] },
		tableRow: {
			content: 'tableCell+',
			attrs: { properties: { default: null } },
			toDOM: () => ['tr', 0],
		},
		tableCell: {
			content: 'paragraph+',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['td', 0],
		},
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
const name: ContentControl = { id: '65', level: 'inline', kind: 'text' };
const named = (text: string) =>
	schema.text(text, [schema.marks.contentControl!.create({ data: JSON.stringify(name) })]);
const plugin = contentControlPlugin();
/** The controls of `slice` as pasted with the selection of `state`. */
const pasted = (
	state: EditorState,
	slice: ReturnType<EditorState['doc']['slice']>,
	plain = false,
) =>
	controlSpans(
		schema.node(
			'doc',
			null,
			plugin.props
				.transformPasted!.call(
					plugin,
					slice,
					{ state, dragging: null } as unknown as EditorView,
					plain,
				)
				.content.content.flatMap((node) =>
					node.isInline ? [schema.node('paragraph', null, [node])] : [node],
				),
		),
	);

describe('content controls pasted in the editor', () => {
	it('keeps the id of a control that text copied from it is pasted back into', () => {
		const doc = schema.node('doc', null, [
			schema.node('paragraph', null, [schema.text('Name: '), named('Ann')]),
		]);
		const state = EditorState.create({ schema, doc });
		// "An", copied from inside the control, with its mark.
		const slice = doc.slice(7, 9);
		const inside = state.apply(state.tr.setSelection(TextSelection.create(doc, 8)));
		expect(pasted(inside, slice).map((span) => span.control.id)).toEqual(['65']);
		// Plain text takes the marks of where it lands: never a copy of a control.
		const before = state.apply(state.tr.setSelection(TextSelection.create(doc, 2)));
		expect(pasted(before, slice, true).map((span) => span.control.id)).toEqual(['65']);
		// Pasted as a copy outside the control, it is a control of its own.
		const [copy] = pasted(before, slice);
		expect(copy!.control.id).not.toBe('65');
	});

	it('joins a control only where typed text would, and not from a drag', () => {
		const doc = schema.node('doc', null, [
			schema.node('paragraph', null, [schema.text('Name: '), named('Ann')]),
		]);
		const state = EditorState.create({ schema, doc });
		const slice = doc.slice(7, 9);
		const at = (pos: number) => state.apply(state.tr.setSelection(TextSelection.create(doc, pos)));
		// Right before the control, after plain text, typed text is not part of it: a copy.
		expect(pasted(at(7), slice)[0]!.control.id).not.toBe('65');
		// At its end, typed text is part of it.
		expect(pasted(at(10), slice).map((span) => span.control.id)).toEqual(['65']);
		// A copy dragged out of the control: the selection is where the drag came from.
		const dragged = plugin.props.transformPasted!.call(
			plugin,
			slice,
			{ state: at(8), dragging: { slice, move: false } } as unknown as EditorView,
			false,
		);
		expect(
			controlSpans(schema.node('doc', null, [schema.node('paragraph', null, dragged.content)]))[0]!
				.control.id,
		).not.toBe('65');
	});

	it('gives a pasted copy of a row control a new id', () => {
		const item: ContentControl = { id: '66', level: 'row', kind: 'richText' };
		const row = schema.node(
			'tableRow',
			{ properties: { heightTwips: 400, contentControl: item } },
			[schema.node('tableCell', null, [schema.node('paragraph', null, [schema.text('Line')])])],
		);
		const doc = schema.node('doc', null, [schema.node('table', null, [row])]);
		const state = EditorState.create({ schema, doc });
		const slice = doc.slice(0, doc.content.size);
		const result = plugin.props.transformPasted!.call(
			plugin,
			slice,
			{ state, dragging: null } as unknown as EditorView,
			false,
		);
		const properties = result.content.firstChild!.firstChild!.attrs.properties;
		expect(properties.heightTwips).toBe(400);
		expect(properties.contentControl.id).not.toBe('66');
		expect(properties.contentControl.id).toMatch(/^\d+$/);
	});
});

describe('freshControlId', () => {
	it('stays within the signed 32-bit integers Word uses for control ids', () => {
		const random = vi.spyOn(Math, 'random');
		try {
			random.mockReturnValue(1 - 2 ** -53);
			expect(Number(freshControlId(new Set()))).toBe(2 ** 31 - 1);
			random.mockReturnValue(0);
			expect(freshControlId(new Set())).toBe('1');
		} finally {
			random.mockRestore();
		}
	});
});
