import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { createDocument, type Paragraph } from '../model';
import { sectionsOf } from '../section-layout';
import { sectionPartsJson } from './header-footer-history';
import { hasSectionPartRevisions, resolveSectionPartRevisions } from './review-section-parts';
import { EditorState } from 'prosemirror-state';
import { listRevisions } from '../revision-commands';
import { history, undo, redo } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { acceptAllChanges, rejectAllChanges, hasAnyChange } from './review-commands';
import { markSpecs } from './schema-marks';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+', attrs: { sectionParts: { default: null } } },
		paragraph: { content: 'text*' },
		text: {},
	},
	marks: markSpecs,
});
const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function source() {
	const paragraph: Paragraph = {
		type: 'paragraph',
		id: 'header',
		align: 'right',
		formatRevision: {
			id: 'p',
			kind: 'paragraphChange',
			author: 'Ada',
			previousParagraphPropertiesXml: `<w:pPr xmlns:w="${ns}"><w:jc w:val="center"/></w:pPr>`,
		},
		runs: [
			{ text: 'Added', revision: { id: 'insert', kind: 'insert', author: 'Ada' } },
			{ text: 'Removed', revision: { id: 'delete', kind: 'delete', author: 'Bob' } },
			{
				text: 'Format',
				bold: true,
				formatRevision: {
					id: 'f',
					kind: 'formatChange',
					author: 'Ada',
					previousRunPropertiesXml: `<w:rPr xmlns:w="${ns}"><w:b w:val="0"/></w:rPr>`,
				},
			},
		],
	};
	const [section] = sectionsOf(createDocument());
	return [
		{
			...section!,
			endsAtBlockId: 'body',
			headers: { default: { partName: 'word/header1.xml', blocks: [paragraph] } },
			footers: {
				default: {
					partName: 'word/footer1.xml',
					blocks: [
						{
							type: 'table' as const,
							id: 'table',
							rows: [[{ paragraphs: [{ ...paragraph, id: 'footer' }] }]],
						},
					],
				},
			},
		},
	];
}
for (const mode of ['accept', 'reject'] as const)
	it(`resolves ${mode} in stored headers and footer tables without changing the source`, () => {
		const sections = source();
		const initial = sectionPartsJson(sections);
		const doc = schema.node(
			'doc',
			{ sectionParts: initial },
			schema.node('paragraph', null, schema.text('Body')),
		);
		expect(hasSectionPartRevisions(doc)).toBe(true);
		const step = resolveSectionPartRevisions(doc, mode)!;
		const state = EditorState.create({ doc });
		const resolved = state.apply(state.tr.step(step)).doc;
		const parts = JSON.parse(resolved.attrs.sectionParts) as typeof sections;
		expect(listRevisions({ ...createDocument(), blocks: [], sections: parts })).toHaveLength(0);
		const header = parts[0]!.headers.default.blocks[0]! as Paragraph;
		expect(header.runs.map((run) => run.text).join('')).toBe(
			mode === 'accept' ? 'AddedFormat' : 'RemovedFormat',
		);
		expect(header.align).toBe(mode === 'accept' ? 'right' : 'center');
		expect(header.runs.at(-1)!.bold).toBe(mode === 'accept');
		expect(resolved.textContent).toBe('Body');
		expect(sectionPartsJson(sections)).toBe(initial);
	});

it('preserves body and story history when a stored paragraph boundary cannot be removed', () => {
	const sections = source();
	const paragraph = sections[0]!.headers.default.blocks[0]! as Paragraph;
	paragraph.markRevision = { id: 'boundary', kind: 'insert', author: 'Ada' };
	const editor = {
		state: EditorState.create({
			doc: schema.node(
				'doc',
				{ sectionParts: sectionPartsJson(sections) },
				schema.node(
					'paragraph',
					null,
					schema.text('Body', [schema.marks.insertion!.create({ id: 'body', author: 'Ada' })]),
				),
			),
			plugins: [history()],
		}),
		editable: true,
		dispatch(this: { state: EditorState }, tr: import('prosemirror-state').Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	} as unknown as EditorView;
	const initial = editor.state.doc;
	expect(() => rejectAllChanges(editor)).toThrow('no following paragraph');
	expect(editor.state.doc).toBe(initial);
	expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(false);
});

for (const mode of ['accept', 'reject'] as const)
	it(`enables ${mode} for story-only revisions with one undo operation`, () => {
		const editor = {
			state: EditorState.create({
				doc: schema.node(
					'doc',
					{ sectionParts: sectionPartsJson(source()) },
					schema.node('paragraph', null, schema.text('Body')),
				),
				plugins: [history()],
			}),
			editable: true,
			dispatch(this: { state: EditorState }, tr: import('prosemirror-state').Transaction) {
				this.state = this.state.apply(tr);
			},
			focus() {},
		} as unknown as EditorView;
		const initial = editor.state.doc;
		expect(hasAnyChange(editor)).toBe(true);
		expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(editor)).toBe(true);
		expect(hasAnyChange(editor)).toBe(false);
		const resolved = editor.state.doc;
		expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(true);
		expect(editor.state.doc.eq(initial)).toBe(true);
		expect(redo(editor.state, editor.dispatch.bind(editor))).toBe(true);
		expect(editor.state.doc.eq(resolved)).toBe(true);
	});
