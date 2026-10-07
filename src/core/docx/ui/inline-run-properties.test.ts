// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { DOMParser, DOMSerializer, Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { TextRun } from '../model';
import {
	hardBreakNodeSpec,
	pageBreakNodeSpec,
	noteReferenceNodeSpec,
	fieldMarkerNodeSpec,
} from './break-note-schema';
import { imageNodeSpec } from './inline-content-schema';
import { markSpecs } from './schema-marks';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';
import { applyRunFormattingPatch } from './run-format-command';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', parseDOM: [{ tag: 'p' }], toDOM: () => ['p', 0] },
		text: { group: 'inline' },
		hardBreak: hardBreakNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
		image: imageNodeSpec,
	},
	marks: markSpecs,
});

for (const atom of [
	{ text: '\n' },
	{ text: '', break: 'column' },
	{ text: '', noteReference: { kind: 'endnote', id: '2' } },
	{ text: '', fieldCode: 'DATE' },
	{
		text: '',
		image: {
			relId: 'rId1',
			partName: 'word/media/a.png',
			contentType: 'image/png',
			widthPx: 10,
			heightPx: 20,
		},
	},
] satisfies TextRun[])
	it(`copies effective ${atom.text ? 'line-break' : Object.keys(atom)[1]} properties independently of its imported basis`, () => {
		const run = { ...atom, bold: true, language: 'en-GB', fontSize: 12 };
		const doc = schema.node(
			'doc',
			null,
			schema.node('paragraph', null, runToInlineNodes(run, schema)),
		);
		let state = EditorState.create({ doc, selection: TextSelection.create(doc, 1, 2) });
		applyRunFormattingPatch({ bold: undefined, fontSize: 18 })(state, (tr) => {
			state = state.apply(tr);
		});
		const original = inlineNodeRun(state.doc.firstChild!.firstChild!)!;
		const container = document.createElement('div');
		container.append(DOMSerializer.fromSchema(schema).serializeFragment(state.doc.content));
		const pasted = DOMParser.fromSchema(schema).parse(container);
		expect(inlineNodeRun(pasted.firstChild!.firstChild!)).toEqual(original);
		expect(original).toMatchObject({ ...atom, language: 'en-GB', fontSize: 18 });
		expect(original).not.toHaveProperty('bold');
	});
