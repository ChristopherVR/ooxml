import { describe, expect, it } from 'vitest';
import { Schema, Slice } from 'prosemirror-model';
import type { ContentControl } from '../model';
import { markSpecs } from './schema-marks';
import { runToInlineNodes } from './run-adapter';
import { controlSourcesOf, pastedControlSources } from './content-control-sources';
import { controlSpans } from './content-control-spans';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0] },
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const own = `<w:sdtPr xmlns:w="${W}"><w:id w:val="11"/><w:dataBinding w:xpath="/a" w:storeItemID="{1}"/></w:sdtPr>`;
const paragraph = (text: string, control: ContentControl) =>
	schema.node('paragraph', null, runToInlineNodes({ text, contentControl: control }, schema));
const sources = (slice: Slice) =>
	controlSpans(schema.node('doc', null, slice.content)).map((span) => span.control.sourceXml);

describe('verbatim properties of pasted controls', () => {
	const doc = schema.node('doc', null, [
		paragraph('Mine', { id: '11', level: 'inline', kind: 'text', sourceXml: own }),
	]);

	it('keeps the verbatim properties of a copy of one of the document own controls', () => {
		const slice = new Slice(
			schema.node('doc', null, [
				paragraph('Copy', { id: '11', level: 'inline', kind: 'text', sourceXml: own }),
			]).content,
			0,
			0,
		);
		expect(pastedControlSources(slice, doc)).toBe(slice);
		expect(sources(pastedControlSources(slice, doc))).toEqual([own]);
	});

	it('takes off verbatim properties the document does not hold, keeping the modeled fields', () => {
		const foreign = `<w:sdtPr xmlns:w="${W}"><w:id w:val="12"/><w:alias w:val="Theirs"/></w:sdtPr>`;
		const slice = new Slice(
			schema.node('doc', null, [
				paragraph('Pasted', {
					id: '12',
					level: 'inline',
					kind: 'text',
					alias: 'Theirs',
					sourceXml: foreign,
				}),
			]).content,
			0,
			0,
		);
		const pasted = pastedControlSources(slice, doc);
		expect(sources(pasted)).toEqual([undefined]);
		const span = controlSpans(schema.node('doc', null, pasted.content))[0]!;
		expect(span.control).toEqual({ id: '12', level: 'inline', kind: 'text', alias: 'Theirs' });
		expect(schema.node('doc', null, pasted.content).textContent).toBe('Pasted');
	});

	it('keeps the verbatim properties a copy, cut or drag in the editor took, as for a move', () => {
		const moved = new Slice(
			schema.node('doc', null, [
				paragraph('Moved', { id: '11', level: 'inline', kind: 'text', sourceXml: own }),
			]).content,
			0,
			0,
		);
		const empty = schema.node('doc', null, [schema.node('paragraph')]);
		expect(controlSourcesOf(moved)).toEqual([own]);
		// Cut, the control is no longer in the document, but the editor took its properties.
		expect(sources(pastedControlSources(moved, empty, new Set(controlSourcesOf(moved))))).toEqual([
			own,
		]);
		expect(sources(pastedControlSources(moved, empty))).toEqual([undefined]);
	});

	it('takes only the values of a pasted control that are of the kind its fields hold', () => {
		// A page can give any JSON: an object where text is drawn or saved (one whose toString is
		// null cannot even become text) is not taken, and the rest of the control is; without an
		// id and a kind that are text, it is no control and its text is pasted alone.
		const crafted = { toString: null };
		const marked = (text: string, data: unknown) =>
			schema.text(text, [schema.marks.contentControl!.create({ data: JSON.stringify(data) })]);
		const slice = new Slice(
			schema.node('doc', null, [
				schema.node('paragraph', null, [
					marked('Pasted', {
						id: '13',
						level: 'inline',
						kind: 'comboBox',
						alias: crafted,
						lock: crafted,
						showingPlaceholder: 'yes',
						checkedGlyph: { code: crafted },
						items: [
							{ display: crafted, value: 'x' },
							{ display: 7, value: 7 },
							{ display: crafted },
						],
					}),
					marked('Plain', { id: crafted, level: 'inline', kind: 'text' }),
				]),
			]).content,
			0,
			0,
		);
		const pasted = schema.node('doc', null, pastedControlSources(slice, doc).content);
		expect(controlSpans(pasted).map((span) => span.control)).toEqual([
			{
				id: '13',
				level: 'inline',
				kind: 'comboBox',
				items: [
					{ display: 'x', value: 'x' },
					{ display: '7', value: '7' },
				],
			},
		]);
		expect(pasted.textContent).toBe('PastedPlain');
	});

	it('takes no control from JSON that nests deeper than a control does', () => {
		// Turning it back into JSON could overflow the stack, in the paste or in every later save.
		const deep = '['.repeat(20000) + ']'.repeat(20000);
		const data = `{"id":"14","level":"inline","kind":"text","extra":${deep}}`;
		const slice = new Slice(
			schema.node('doc', null, [
				schema.node('paragraph', null, [
					schema.text('Deep', [schema.marks.contentControl!.create({ data })]),
				]),
			]).content,
			0,
			0,
		);
		const pasted = schema.node('doc', null, pastedControlSources(slice, doc).content);
		expect(controlSpans(pasted)).toEqual([]);
		expect(pasted.textContent).toBe('Deep');
	});
});
