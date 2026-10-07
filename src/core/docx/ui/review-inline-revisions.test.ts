import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { Schema, type Node } from 'prosemirror-model';
import { loadDocx } from '../index';
import { EditorState, type Transaction } from 'prosemirror-state';
import { history, undo, redo } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import type { TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { imageNodeSpec } from './inline-content-schema';
import { pageBreakNodeSpec, noteReferenceNodeSpec, fieldMarkerNodeSpec } from './break-note-schema';
import { runToInlineNodes, inlineNodeRun } from './run-adapter';
import {
	acceptAllChanges,
	rejectAllChanges,
	collectRevisionRanges,
	acceptRevisionRange,
	rejectRevisionRange,
} from './review-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		image: imageNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
		equation: {
			inline: true,
			group: 'inline',
			attrs: { omml: {}, display: {}, format: { default: null } },
		},
	},
	marks: markSpecs,
});
const atoms: TextRun[] = [
	{ text: '', break: 'page' },
	{ text: '', noteReference: { kind: 'footnote', id: '1' } },
	{
		text: '',
		image: {
			relId: 'rId1',
			partName: 'word/media/image.png',
			contentType: 'image/png',
			widthPx: 20,
			heightPx: 20,
		},
	},
	{ text: '', fieldCode: ' PAGE ' },
	{ text: '', equation: { omml: '<m:oMath/>', display: false } },
];
function view(runs: TextRun[], editable = true, document?: Node): EditorView {
	const editor = {
		editable,
		state: EditorState.create({
			doc:
				document ??
				schema.node(
					'doc',
					null,
					schema.node(
						'paragraph',
						null,
						runs.flatMap((run) => runToInlineNodes(run, schema)),
					),
				),
			plugins: [history()],
		}),
		dispatch(tr: Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	};
	return editor as unknown as EditorView;
}
for (const [index, atom] of atoms.entries())
	for (const mode of ['accept', 'reject'] as const)
		it(`${mode}s inline object ${index} text revisions in one undoable transaction`, () => {
			const properties = { bold: false, sourceRunPropertiesXml: '<w:rPr><w:b w:val="0"/></w:rPr>' };
			const editor = view([
				{
					...atom,
					...properties,
					revision: { kind: 'insert', id: 'i', author: 'Ada' },
				},
				{ text: 'Keep' },
				{
					...atom,
					...properties,
					revision: { kind: 'delete', id: 'd', author: 'Ada' },
				},
			]);
			const source = editor.state.doc;
			expect(collectRevisionRanges(source)).toHaveLength(2);
			expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(editor)).toBe(true);
			const runs = editor.state.doc.firstChild!.content.content.map(inlineNodeRun);
			expect(runs).toEqual(
				mode === 'accept'
					? [{ ...atom, ...properties }, { text: 'Keep' }]
					: [{ text: 'Keep' }, { ...atom, ...properties }],
			);
			expect(collectRevisionRanges(editor.state.doc)).toEqual([]);
			const resolved = editor.state.doc;
			expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(true);
			expect(editor.state.doc.eq(source)).toBe(true);
			expect(redo(editor.state, editor.dispatch.bind(editor))).toBe(true);
			expect(editor.state.doc.eq(resolved)).toBe(true);
		});

for (const mode of ['accept', 'reject'] as const)
	it(`resolves a linked atom/text move on ${mode} while retaining another author's reused id`, () => {
		const atom = atoms[0]!;
		const editor = view([
			{ ...atom, revision: { kind: 'moveFrom', id: '1', author: 'Ada', move: { name: 'move' } } },
			{ text: 'Keep' },
			{
				text: 'Moved',
				revision: { kind: 'moveTo', id: '2', author: 'Ada', move: { name: 'move' } },
			},
			{ ...atom, revision: { kind: 'insert', id: '1', author: 'Bob' } },
		]);
		const range = collectRevisionRanges(editor.state.doc)[0]!;
		(mode === 'accept' ? acceptRevisionRange : rejectRevisionRange)(editor, range);
		expect(collectRevisionRanges(editor.state.doc)).toMatchObject([
			{ author: 'Bob', kind: 'insert', id: '1' },
		]);
		expect(editor.state.doc.textContent).toBe(mode === 'accept' ? 'KeepMoved' : 'Keep');
	});

it('retains independent formatting history when resolving an atom text revision', () => {
	const formatRevision = {
		kind: 'formatChange' as const,
		id: 'format',
		author: 'Bob',
		previousRunPropertiesXml: '<w:rPr/>',
	};
	const editor = view([
		{ ...atoms[0]!, formatRevision, revision: { id: 'text', kind: 'insert', author: 'Ada' } },
	]);
	const range = collectRevisionRanges(editor.state.doc).find((item) => item.kind === 'insert')!;
	acceptRevisionRange(editor, range);
	expect(inlineNodeRun(editor.state.doc.firstChild!.firstChild!)).toMatchObject({
		break: 'page',
		formatRevision,
	});
	expect(inlineNodeRun(editor.state.doc.firstChild!.firstChild!)!.revision).toBeUndefined();
});

it('blocks read-only atom resolution and ignores an already resolved range', () => {
	const editor = view(
		[{ ...atoms[0]!, revision: { kind: 'insert', id: 'i', author: 'Ada' } }],
		false,
	);
	const source = editor.state.doc;
	const range = collectRevisionRanges(source)[0]!;
	acceptRevisionRange(editor, range);
	expect(editor.state.doc).toBe(source);
	expect(rejectAllChanges(editor)).toBe(false);
	const writable = view([{ ...atoms[0]!, revision: { kind: 'insert', id: 'i', author: 'Ada' } }]);
	acceptRevisionRange(writable, range);
	const resolved = writable.state.doc;
	rejectRevisionRange(writable, range);
	expect(writable.state.doc).toBe(resolved);
});

function body(model: import('../model').DocumentModel): Node {
	return schema.node(
		'doc',
		null,
		model.blocks.map((block) => {
			if (block.type !== 'paragraph') throw new Error('Expected a native paragraph');
			return schema.node(
				'paragraph',
				null,
				block.runs.flatMap((run) => runToInlineNodes(run, schema)),
			);
		}),
	);
}
function contents(doc: Node) {
	return doc.content.content.map((paragraph) => {
		const result: unknown[] = [];
		paragraph.forEach((node) => {
			const run = inlineNodeRun(node)!;
			if (run.image) result.push({ image: [run.image.widthPx, run.image.heightPx] });
			else if (run.break) result.push({ break: run.break });
			else if (run.noteReference) result.push({ note: run.noteReference });
			else if (run.text) {
				if (typeof result.at(-1) === 'string') result[result.length - 1] += run.text;
				else result.push(run.text);
			}
		});
		return result;
	});
}
for (const name of [
	'picture-insert',
	'picture-delete',
	'note-insert',
	'note-delete',
	'break-delete',
])
	for (const mode of ['accept', 'reject'] as const)
		it(`${mode}s native ${name} body content like desktop Word`, async () => {
			const load = async (suffix: string) =>
				(
					await loadDocx(
						await readFile(
							new URL(`../__fixtures__/review-inline/${name}-${suffix}.docx`, import.meta.url),
						),
					)
				).model;
			const source = await load('tracked');
			const original = structuredClone(source);
			const editor = view([], true, body(source));
			expect(collectRevisionRanges(editor.state.doc)).toHaveLength(1);
			expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(editor)).toBe(true);
			expect(contents(editor.state.doc)).toEqual(
				contents(body(await load(mode === 'accept' ? 'accepted' : 'rejected'))),
			);
			expect(collectRevisionRanges(editor.state.doc)).toEqual([]);
			expect(source).toEqual(original);
		});
