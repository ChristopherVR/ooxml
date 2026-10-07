import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { history, undo, redo } from 'prosemirror-history';
import { loadDocx } from '../parse';
import { listRevisions } from '../revision-commands';
import type { Paragraph, TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { imageNodeSpec } from './inline-content-schema';
import { pageBreakNodeSpec, noteReferenceNodeSpec, fieldMarkerNodeSpec } from './break-note-schema';
import { runToInlineNodes, appendInlineNode, inlineNodeRun } from './run-adapter';
import { trackChangesPlugin, trackChangesPluginKey } from './track-changes-mode';
import { resolveFormattingRange } from './review-formatting';
import { createToggleFormat } from './toggle-format';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+', attrs: { trackFormatting: { default: true } } },
		paragraph: { content: 'inline*', attrs: { id: { default: '' } } },
		text: { group: 'inline' },
		image: imageNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
const normalized = (paragraphs: Paragraph[]) =>
	paragraphs.map((p) =>
		p.runs.map(
			({ sourceRunPropertiesXml: _source, restoredRunPropertiesXml: _restored, ...run }) => run,
		),
	);

for (const name of ['picture', 'note', 'break', 'field'])
	it(`records native ${name} formatting, exports Word-equivalent resolution and retains isolated history`, async () => {
		const fixture = (state: string) =>
			readFile(
				new URL(`../__fixtures__/review-object-formatting/${name}-${state}.docx`, import.meta.url),
			);
		const loaded = await loadDocx(await fixture('before'));
		const paragraphs = loaded.model.blocks as Paragraph[];
		let editor = EditorState.create({
			doc: schema.node(
				'doc',
				null,
				paragraphs.map((p) =>
					schema.node(
						'paragraph',
						{ id: p.id },
						p.runs.flatMap((run) => runToInlineNodes(run, schema)),
					),
				),
			),
			plugins: [
				history(),
				trackChangesPlugin(
					() => 'Ada',
					() => true,
				),
			],
		});
		let position = -1;
		editor.doc.descendants((node, pos) => {
			if (position >= 0 || !node.isInline || node.isText) return;
			if (node.type.name === 'fieldMarker' && node.attrs.kind !== 'code') return;
			position = pos;
		});
		expect(position).toBeGreaterThan(0);
		const initial = editor.doc;
		editor = editor.apply(
			editor.tr.setSelection(TextSelection.create(editor.doc, position, position + 1)),
		);
		createToggleFormat(schema, () => loaded.model)('bold')(editor, (tr) => {
			editor = editor.apply(tr);
		});
		const changed = editor.doc;
		expect(
			inlineNodeRun(changed.nodeAt(position)!)!.formatRevision?.previousRunPropertiesXml,
		).toContain('rPr');
		for (const mode of ['accept', 'reject'] as const) {
			const tr = editor.tr;
			resolveFormattingRange(tr, position, position + 1, mode);
			const state = editor.apply(tr.setMeta(trackChangesPluginKey, { tracked: true }));
			const model = {
				...loaded.model,
				blocks: paragraphs.map((p, index) => {
					const runs: TextRun[] = [];
					state.doc.child(index).forEach((node) => appendInlineNode(runs, node));
					return { ...p, runs };
				}),
			};
			const reopened = await loadDocx(await loaded.save(model));
			const native = await loadDocx(await fixture(mode === 'accept' ? 'accepted' : 'rejected'));
			expect(listRevisions(reopened.model)).toHaveLength(0);
			expect(normalized(reopened.model.blocks as Paragraph[])).toEqual(
				normalized(native.model.blocks as Paragraph[]),
			);
		}
		expect(
			undo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.eq(initial)).toBe(true);
		expect(
			redo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.eq(changed)).toBe(true);
		const pending = inlineNodeRun(changed.nodeAt(position)!)!.formatRevision;
		createToggleFormat(schema, () => loaded.model)('italic')(editor, (tr) => {
			editor = editor.apply(tr);
		});
		expect(inlineNodeRun(editor.doc.nodeAt(position)!)!.formatRevision).toEqual(pending);
		createToggleFormat(schema, () => loaded.model)('bold')(editor, (tr) => {
			editor = editor.apply(tr);
		});
		expect(inlineNodeRun(editor.doc.nodeAt(position)!)!.formatRevision).toEqual(pending);
		createToggleFormat(schema, () => loaded.model)('italic')(editor, (tr) => {
			editor = editor.apply(tr);
		});
		expect(inlineNodeRun(editor.doc.nodeAt(position)!)!.formatRevision).toBeUndefined();
	});

for (const mode of ['accept', 'reject'] as const)
	it(`keeps independently tracked note insertion when resolving newly recorded formatting: ${mode}`, () => {
		const run: TextRun = {
			text: '',
			noteReference: { kind: 'endnote', id: '7' },
			revision: { id: 'insert', kind: 'insert', author: 'Bob' },
			commentIds: ['comment'],
			language: 'en-US',
		};
		let state = EditorState.create({
			doc: schema.node('doc', null, schema.node('paragraph', null, runToInlineNodes(run, schema))),
			plugins: [
				trackChangesPlugin(
					() => 'Ada',
					() => true,
				),
			],
		});
		state = state.apply(state.tr.addMark(1, 2, schema.marks.bold!.create()));
		expect(inlineNodeRun(state.doc.nodeAt(1)!)!.formatRevision?.author).toBe('Ada');
		const tr = state.tr;
		resolveFormattingRange(tr, 1, 2, mode);
		state = state.apply(tr.setMeta(trackChangesPluginKey, { tracked: true }));
		const resolved = inlineNodeRun(state.doc.nodeAt(1)!)!;
		expect(resolved.revision).toEqual(run.revision);
		expect(resolved.commentIds).toEqual(run.commentIds);
		expect(resolved.noteReference).toEqual(run.noteReference);
		expect(resolved.language).toBe('en-US');
		expect(resolved.formatRevision).toBeUndefined();
		expect(resolved.bold).toBe(mode === 'accept' ? true : undefined);
	});
