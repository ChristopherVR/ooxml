import { describe, expect, it, vi } from 'vitest';
import { Fragment, Schema, Slice, type Node as ProseMirrorNode } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { ContentControl } from '../model';
import { markSpecs } from './schema-marks';
import { contentControlPlugin } from './content-control-guard';
import { freshControlId } from '../content-control-model';
import { controlFromData, controlFromValue, controlSpans } from './content-control-spans';

const schema = new Schema({
	nodes: {
		doc: { content: 'block+' },
		paragraph: {
			group: 'block',
			content: 'inline*',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['p', 0],
		},
		table: {
			group: 'block',
			content: 'tableRow+',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['table', ['tbody', 0]],
		},
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

	describe('verbatim properties', () => {
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const properties = (id: string, store: string) =>
			`<w:sdtPr xmlns:w="${W}"><w:id w:val="${id}"/><w:dataBinding w:xpath="/a" w:storeItemID="${store}"/></w:sdtPr>`;
		const end = (run: string) =>
			`<w:sdtEndPr xmlns:w="${W}"><w:rPr><w:${run}/></w:rPr></w:sdtEndPr>`;
		const table = (control: ContentControl) =>
			schema.node('table', null, [
				schema.node('tableRow', null, [
					schema.node('tableCell', { contentControl: control }, [
						schema.node('paragraph', null, [schema.text('Cell')]),
					]),
				]),
			]);
		const marked = (control: ContentControl) =>
			schema.node('paragraph', null, [
				schema.text('Text', [
					schema.marks.contentControl!.create({ data: JSON.stringify(control) }),
				]),
			]);
		const view = (doc: ProseMirrorNode) =>
			({ state: EditorState.create({ schema, doc }), dragging: null }) as unknown as EditorView;
		const empty = schema.node('doc', null, [schema.node('paragraph')]);
		/** The control of `block` pasted into `doc` through `editor`'s content control plugin. */
		const paste = (
			editor: ReturnType<typeof contentControlPlugin>,
			doc: ProseMirrorNode,
			block: ProseMirrorNode,
		) => {
			const slice = new Slice(Fragment.from(block), 0, 0);
			const result = editor.props.transformPasted!.call(editor, slice, view(doc), false);
			return controlSpans(schema.node('doc', null, result.content))[0]!.control;
		};

		it('keeps those of a pasted cell control only when they come from the editor', () => {
			const mine: ContentControl = {
				id: '31',
				level: 'cell',
				kind: 'richText',
				sourceXml: properties('31', '{1}'),
				endSourceXml: end('b'),
			};
			// A table from another page: properties that bind its cell to that page's data.
			const theirs: ContentControl = {
				id: '32',
				level: 'cell',
				kind: 'richText',
				alias: 'Theirs',
				sourceXml: properties('32', '{2}'),
				endSourceXml: end('i'),
			};
			const doc = schema.node('doc', null, [table(mine)]);
			const editor = contentControlPlugin();
			// A copy of one of the document's own cells keeps them, under a new id.
			const copy = paste(editor, doc, table(mine));
			expect(copy).toMatchObject({ sourceXml: mine.sourceXml, endSourceXml: mine.endSourceXml });
			expect(copy.id).not.toBe('31');
			// The other page's are taken off, and its modeled fields kept.
			expect(paste(editor, doc, table(theirs))).toEqual({
				id: '32',
				level: 'cell',
				kind: 'richText',
				alias: 'Theirs',
			});
			// Cut, a cell is no longer in the document: it keeps them when the editor took them.
			expect(paste(editor, empty, table(mine))).toEqual({
				id: '31',
				level: 'cell',
				kind: 'richText',
			});
			editor.props.transformCopied!.call(
				editor,
				new Slice(Fragment.from(table(mine)), 0, 0),
				view(doc),
			);
			expect(paste(editor, empty, table(mine))).toEqual(mine);
		});

		it('keeps those of a pasted row, table or block control only when they come from the editor', () => {
			// The slice context of pasted HTML (data-pm-slice) gives the nodes around what it holds
			// any attributes, so the controls of rows, tables and paragraphs come with it too.
			const control = (id: string, level: ContentControl['level'], store: string, run: string) =>
				({
					id,
					level,
					kind: 'richText',
					sourceXml: properties(id, store),
					endSourceXml: end(run),
				}) satisfies ContentControl;
			const cell = () =>
				schema.node('tableCell', null, [schema.node('paragraph', null, [schema.text('Cell')])]);
			const holding = (level: 'row' | 'table' | 'paragraph', held: ContentControl) =>
				level === 'paragraph'
					? schema.node('paragraph', { contentControl: held }, [schema.text('Text')])
					: schema.node('table', level === 'table' ? { contentControl: held } : null, [
							schema.node(
								'tableRow',
								level === 'row' ? { properties: { heightTwips: 400, contentControl: held } } : null,
								[cell()],
							),
						]);
			const controlOf = (node: ProseMirrorNode, level: 'row' | 'table' | 'paragraph') =>
				level === 'row'
					? node.firstChild!.attrs.properties.contentControl
					: node.attrs.contentControl;
			const levels = { row: 'row', table: 'block', paragraph: 'block' } as const;
			for (const level of ['row', 'table', 'paragraph'] as const) {
				const mine = control('51', levels[level], '{1}', 'b');
				const theirs = control('52', levels[level], '{2}', 'i');
				const doc = schema.node('doc', null, [holding(level, mine)]);
				const editor = contentControlPlugin();
				const pastedNode = (into: ProseMirrorNode, block: ProseMirrorNode) =>
					editor.props.transformPasted!.call(
						editor,
						new Slice(Fragment.from(block), 0, 0),
						view(into),
						false,
					).content.firstChild!;
				expect(controlOf(pastedNode(doc, holding(level, mine)), level), level).toMatchObject({
					sourceXml: mine.sourceXml,
					endSourceXml: mine.endSourceXml,
				});
				expect(controlOf(pastedNode(doc, holding(level, theirs)), level), level).toEqual({
					id: '52',
					level: levels[level],
					kind: 'richText',
				});
				editor.props.transformCopied!.call(
					editor,
					new Slice(Fragment.from(holding(level, mine)), 0, 0),
					view(doc),
				);
				expect(controlOf(pastedNode(empty, holding(level, mine)), level), level).toEqual(mine);
			}
		});

		it('takes no control from a pasted node that holds what is not one', () => {
			// What the slice context gives a node is not read through the parse rules: only an object
			// with an id and a kind is a control, at the level of where it sits.
			const nodes = schema.node('doc', null, [
				schema.node('paragraph', { contentControl: [1] }, [schema.text('Text')]),
				schema.node('table', { contentControl: 'table' }, [
					schema.node(
						'tableRow',
						{ properties: { heightTwips: 400, contentControl: { id: '7' } } },
						[
							schema.node('tableCell', { contentControl: 5 }, [schema.node('paragraph')]),
							schema.node(
								'tableCell',
								{ contentControl: { id: '8', kind: 'text', level: 'inline' } },
								[schema.node('paragraph')],
							),
						],
					),
				]),
			]);
			const editor = contentControlPlugin();
			const result = editor.props.transformPasted!.call(
				editor,
				new Slice(nodes.content, 0, 0),
				view(empty),
				false,
			);
			const [paragraph, table] = [result.content.child(0), result.content.child(1)];
			expect(paragraph.attrs.contentControl).toBeNull();
			expect(table.attrs.contentControl).toBeNull();
			expect(table.firstChild!.attrs.properties).toEqual({ heightTwips: 400 });
			expect(table.firstChild!.child(0).attrs.contentControl).toBeNull();
			expect(table.firstChild!.child(1).attrs.contentControl).toEqual({
				id: '8',
				kind: 'text',
				level: 'cell',
			});
		});

		it('takes only the values of a pasted control that are of the kind its fields hold', () => {
			// A page can give any JSON: an object where text is drawn or saved (one whose toString is
			// null cannot even become text) is not taken, and the rest of the control is.
			const crafted = { toString: null };
			const fields = {
				alias: crafted,
				tag: crafted,
				lock: crafted,
				dateFormat: crafted,
				showingPlaceholder: 'yes',
				checkedGlyph: { code: crafted },
				items: [
					{ display: crafted, value: 'x' },
					{ display: 'Shown', value: crafted },
					{ display: 7, value: 7 },
					{ display: crafted },
					{ display: 'Fine', value: 'fine' },
				],
			};
			const taken = {
				items: [
					{ display: 'x', value: 'x' },
					{ display: 'Shown', value: 'Shown' },
					{ display: '7', value: '7' },
					{ display: 'Fine', value: 'fine' },
				],
			};
			const editor = contentControlPlugin();
			const cell = { id: '80', level: 'cell', kind: 'dropDownList', ...fields };
			expect(paste(editor, empty, table(cell as unknown as ContentControl))).toEqual({
				id: '80',
				level: 'cell',
				kind: 'dropDownList',
				...taken,
			});
			const inline = { id: '81', level: 'inline', kind: 'comboBox', ...fields };
			expect(controlFromData(JSON.stringify(inline), 'inline')).toEqual({
				id: '81',
				level: 'inline',
				kind: 'comboBox',
				...taken,
			});
			// A well-formed control is taken as it is.
			const fine: ContentControl = { id: '82', level: 'cell', kind: 'checkbox', checked: true };
			expect(controlFromValue(fine, 'cell')).toBe(fine);
		});

		it('keeps the end properties of a pasted inline control only when they come from the editor', () => {
			const mine: ContentControl = {
				id: '33',
				level: 'inline',
				kind: 'text',
				endSourceXml: end('b'),
			};
			const theirs: ContentControl = {
				id: '34',
				level: 'inline',
				kind: 'text',
				endSourceXml: end('i'),
			};
			const doc = schema.node('doc', null, [marked(mine)]);
			const editor = contentControlPlugin();
			expect(paste(editor, doc, marked(mine)).endSourceXml).toBe(mine.endSourceXml);
			expect(paste(editor, doc, marked(theirs))).toEqual({
				id: '34',
				level: 'inline',
				kind: 'text',
			});
			expect(paste(editor, empty, marked(mine)).endSourceXml).toBeUndefined();
			editor.props.transformCopied!.call(
				editor,
				new Slice(Fragment.from(marked(mine)), 0, 0),
				view(doc),
			);
			expect(paste(editor, empty, marked(mine))).toEqual(mine);
		});
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
