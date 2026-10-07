import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { Schema, type Node } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory, history, redo, undo } from 'prosemirror-history';
import { markSpecs } from './schema-marks';
import { paragraphAttrs, paragraphFromAttrs } from './paragraph-attributes';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';
import { trackChangesPlugin, REMOTE_TRANSACTION_META } from './track-changes-mode';
import { trackParagraphBoundary } from './track-paragraph-boundaries';
import {
	acceptAllChanges,
	rejectAllChanges,
	acceptRevisionRange,
	rejectRevisionRange,
	collectRevisionRanges,
} from './review-commands';
import { loadDocx } from '../parse';
import { restartFixture } from '../test-support/restart-fixture';
import { fieldMarkerNodeSpec, noteReferenceNodeSpec } from './break-note-schema';
import { paragraphBoundaryFixture } from '../test-support/paragraph-boundary-fixture';

const attrs = Object.fromEntries(
	Object.entries(paragraphAttrs({ type: 'paragraph', id: '', runs: [] })).map(([key, value]) => [
		key,
		{ default: value },
	]),
);
const schema = new Schema({
	nodes: {
		doc: { content: 'block+', attrs: { sections: { default: null } } },
		paragraph: { content: 'inline*', group: 'block', attrs },
		table: { content: 'paragraph+', group: 'block' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
		noteReference: noteReferenceNodeSpec,
	},
	marks: markSpecs,
});
const paragraph = (id: string, text: string, align: 'left' | 'right' | 'center' = 'center') =>
	schema.node('paragraph', { id, align }, schema.text(text, [schema.marks.bold!.create()]));
function editor(doc: Node) {
	const view = {
		state: EditorState.create({
			doc,
			plugins: [
				history(),
				trackChangesPlugin(
					() => 'Ada',
					() => true,
				),
			],
		}),
		editable: true,
		dispatch(tr: Transaction) {
			view.state = view.state.applyTransaction(tr).state;
		},
		focus() {},
	};
	return view as unknown as EditorView;
}

for (const action of ['split', 'join'] as const)
	for (const mode of ['accept', 'reject'] as const)
		it(`records a plain paragraph ${action}, ${mode}s it and restores editing/review history`, () => {
			const doc = schema.node(
				'doc',
				null,
				action === 'split'
					? paragraph('original', 'Hello')
					: [paragraph('first', 'He'), paragraph('last', 'llo')],
			);
			const view = editor(doc);
			view.dispatch(action === 'split' ? view.state.tr.split(3) : view.state.tr.join(4));
			const pending = view.state.doc;
			expect(pending.childCount).toBe(2);
			expect(pending.textContent).toBe('Hello');
			expect(collectRevisionRanges(pending).map(({ kind, author }) => ({ kind, author }))).toEqual([
				{ kind: 'paragraphMark', author: 'Ada' },
			]);
			expect(pending.firstChild!.attrs.markRevision.kind).toBe(
				action === 'split' ? 'insert' : 'delete',
			);
			if (action === 'split') {
				expect(pending.lastChild!.attrs.id).toBe('original');
				expect(pending.firstChild!.attrs.id).not.toBe('original');
			}
			expect(undo(view.state, view.dispatch.bind(view))).toBe(true);
			expect(view.state.doc.eq(doc)).toBe(true);
			expect(redo(view.state, view.dispatch.bind(view))).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
			view.dispatch(closeHistory(view.state.tr));
			expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(view)).toBe(true);
			const keep = (mode === 'accept') === (action === 'split');
			expect(view.state.doc.childCount).toBe(keep ? 2 : 1);
			expect(collectRevisionRanges(view.state.doc)).toEqual([]);
			if (!keep)
				expect(view.state.doc.firstChild!.attrs).toMatchObject({
					id: action === 'split' ? 'original' : 'last',
					align: 'center',
				});
			expect(undo(view.state, view.dispatch.bind(view))).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
		});

it("removes the author's pending inserted boundary without adding a deletion", () => {
	const doc = schema.node('doc', null, paragraph('original', 'Hello'));
	const view = editor(doc);
	view.dispatch(view.state.tr.split(3));
	view.dispatch(view.state.tr.join(view.state.doc.firstChild!.nodeSize));
	expect(view.state.doc.eq(doc)).toBe(true);
	expect(collectRevisionRanges(view.state.doc)).toEqual([]);
});

it('excludes bookmarked paragraph splits from the guarded boundary recorder', () => {
	const doc = schema.node(
		'doc',
		null,
		schema.node(
			'paragraph',
			{
				id: 'original',
				bookmarks: ['Target'],
			},
			schema.text('Hello'),
		),
	);
	const state = EditorState.create({ doc });
	const tr = state.tr.split(3);
	expect(
		trackParagraphBoundary(tr.steps, state, state.apply(tr), 'Ada', 'date', () => 'id'),
	).toBeNull();
	const view = editor(doc);
	view.dispatch(view.state.tr.split(3));
	expect(
		collectRevisionRanges(view.state.doc).some((range) => range.kind === 'paragraphMark'),
	).toBe(false);
});

for (const action of ['accept', 'reject'] as const)
	it(`retains inline comment anchors, foreign revised text and note reference identity on ${action}ing a supported split`, () => {
		const original = schema.node('paragraph', { id: 'original' }, [
			schema.text('AB', [
				schema.marks.comment!.create({ ids: ['comment-one'] }),
				schema.marks.insertion!.create({ author: 'Bob', id: 'foreign-text' }),
			]),
			schema.node('noteReference', {
				kind: 'footnote',
				id: 'note-7',
				format: JSON.stringify({ commentIds: ['comment-note'], bold: true }),
			}),
			schema.text('CD', [schema.marks.comment!.create({ ids: ['comment-two'] })]),
		]);
		const view = editor(schema.node('doc', null, original));
		view.dispatch(view.state.tr.split(2));
		expect(
			view.state.doc
				.firstChild!.content.append(view.state.doc.lastChild!.content)
				.eq(original.content),
		).toBe(true);
		const boundary = collectRevisionRanges(view.state.doc).find(
			(range) => range.kind === 'paragraphMark',
		)!;
		(action === 'accept' ? acceptRevisionRange : rejectRevisionRange)(view, boundary);
		const content =
			action === 'accept'
				? view.state.doc.firstChild!.content.append(view.state.doc.lastChild!.content)
				: view.state.doc.firstChild!.content;
		expect(content.eq(original.content)).toBe(true);
		expect(
			collectRevisionRanges(view.state.doc).every((range) => range.id === 'foreign-text'),
		).toBe(true);
		const references: Node[] = [];
		view.state.doc.descendants((node) => {
			if (node.type.name === 'noteReference') references.push(node);
		});
		expect(references).toHaveLength(1);
		expect(inlineNodeRun(references[0]!)).toMatchObject({
			noteReference: { kind: 'footnote', id: 'note-7' },
			commentIds: ['comment-note'],
			bold: true,
		});
	});

it('excludes fields, section boundaries, tables, pending foreign boundaries and mixed replacements', () => {
	const simple = schema.text('Hello', [
		schema.marks.field!.create({ instr: 'REF Target', simple: true }),
	]);
	const cases = [
		{ doc: schema.node('doc', null, schema.node('paragraph', { id: 'field' }, simple)), pos: 3 },
		{
			doc: schema.node('doc', { sections: '[{"endsAtBlockId":"original"}]' }, [
				paragraph('original', 'Hello'),
				paragraph('last', 'Tail'),
			]),
			pos: 3,
		},
		{
			doc: schema.node('doc', null, schema.node('table', null, paragraph('original', 'Hello'))),
			pos: 4,
		},
		{
			doc: schema.node(
				'doc',
				null,
				paragraph('original', 'Hello').type.create(
					{ id: 'original', markRevision: { kind: 'insert', author: 'Bob', id: 'prior' } },
					schema.text('Hello'),
				),
			),
			pos: 3,
		},
	];
	for (const { doc, pos } of cases) {
		const state = EditorState.create({ doc });
		const tr = state.tr.split(pos);
		expect(
			trackParagraphBoundary(tr.steps, state, state.apply(tr), 'Ada', 'date', () => 'id'),
		).toBeNull();
	}
	const state = EditorState.create({
		doc: schema.node('doc', null, paragraph('original', 'Hello')),
	});
	const tr = state.tr.split(3).insertText('X', 2);
	expect(
		trackParagraphBoundary(tr.steps, state, state.apply(tr), 'Ada', 'date', () => 'id'),
	).toBeNull();
});

it('leaves remote paragraph splits untracked', () => {
	const view = editor(schema.node('doc', null, paragraph('original', 'Hello')));
	view.dispatch(view.state.tr.split(3).setMeta(REMOTE_TRANSACTION_META, true));
	expect(view.state.doc.childCount).toBe(2);
	expect(collectRevisionRanges(view.state.doc)).toEqual([]);
});

it('excludes joins needing separate paragraph formatting history', () => {
	const state = EditorState.create({
		doc: schema.node('doc', null, [
			paragraph('first', 'He', 'left'),
			paragraph('last', 'llo', 'right'),
		]),
	});
	const tr = state.tr.join(4);
	expect(
		trackParagraphBoundary(tr.steps, state, state.apply(tr), 'Ada', 'date', () => 'id'),
	).toBeNull();
});

it('keeps opaque source paragraph property differences outside the plain join recorder', () => {
	const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	const nodes = ['0', '1'].map((value, index) =>
		schema.node(
			'paragraph',
			{
				id: `p${index}`,
				sourceParagraphPropertiesXml: `<w:pPr xmlns:w="${w}"><w:wordWrap w:val="${value}"/></w:pPr>`,
			},
			schema.text(index ? 'llo' : 'He'),
		),
	);
	const state = EditorState.create({ doc: schema.node('doc', null, nodes) });
	const tr = state.tr.join(4);
	expect(
		trackParagraphBoundary(tr.steps, state, state.apply(tr), 'Ada', 'date', () => 'id'),
	).toBeNull();
});

it.each(['split', 'join'] as const)(
	'saves and reopens a tracked %s with original paragraph formatting intact',
	async (action) => {
		const zip = await JSZip.loadAsync(await restartFixture(undefined));
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const props =
			'<w:pPr><w:jc w:val="center"/><w:spacing w:after="120" data-extra="keep"/><w:keepNext/></w:pPr>';
		const p = (text: string) =>
			`<w:p>${props}<w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body>${action === 'split' ? p('Hello') : p('He') + p('llo')}<w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const nodes = loaded.model.blocks.map((block) => {
			if (block.type !== 'paragraph') throw new Error('Expected paragraph');
			return schema.node(
				'paragraph',
				paragraphAttrs(block),
				block.runs.flatMap((run) => runToInlineNodes(run, schema)),
			);
		});
		const view = editor(schema.node('doc', null, nodes));
		view.dispatch(action === 'split' ? view.state.tr.split(3) : view.state.tr.join(4));
		const blocks = Array.from({ length: view.state.doc.childCount }, (_, index) => {
			const node = view.state.doc.child(index);
			const runs = Array.from({ length: node.childCount }, (_, run) =>
				inlineNodeRun(node.child(run))!,
			);
			return paragraphFromAttrs(node.attrs, String(node.attrs.id), runs);
		});
		const saved = await loaded.save({ ...loaded.model, blocks });
		const reopened = await loadDocx(saved);
		const paragraphs = reopened.model.blocks;
		if (paragraphs[0]?.type !== 'paragraph' || paragraphs[1]?.type !== 'paragraph')
			throw new Error('Expected paragraphs');
		expect(paragraphs[0].markRevision).toMatchObject({
			kind: action === 'split' ? 'insert' : 'delete',
			author: 'Ada',
		});
		for (const block of paragraphs) {
			expect(block).toMatchObject({ align: 'center', spacingAfterTwips: 120, keepNext: true });
			expect(block.type === 'paragraph' && block.runs.every((run) => run.bold)).toBe(true);
		}
		if (action === 'split') {
			const imported = paragraphs.map((block) => {
				if (block.type !== 'paragraph') throw new Error('Expected paragraph');
				return schema.node(
					'paragraph',
					paragraphAttrs(block),
					block.runs.flatMap((run) => runToInlineNodes(run, schema)),
				);
			});
			const joined = editor(schema.node('doc', null, imported));
			joined.dispatch(joined.state.tr.join(joined.state.doc.firstChild!.nodeSize));
			expect(joined.state.doc.childCount).toBe(1);
			expect(collectRevisionRanges(joined.state.doc)).toEqual([]);
		}
		expect(
			await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string'),
		).toContain('data-extra="keep"');
	},
);

it.each(['split', 'join'] as const)(
	'matches native Word plain %s paragraph-mark placement and formatting',
	async (action) => {
		const source = await loadDocx(await paragraphBoundaryFixture(action));
		const nodes = source.model.blocks.map((block) => {
			if (block.type !== 'paragraph') throw new Error('Expected paragraph');
			return schema.node(
				'paragraph',
				paragraphAttrs(block),
				block.runs.flatMap((run) => runToInlineNodes(run, schema)),
			);
		});
		const view = editor(schema.node('doc', null, nodes));
		view.dispatch(action === 'split' ? view.state.tr.split(3) : view.state.tr.join(4));
		const blocks = Array.from({ length: view.state.doc.childCount }, (_, index) => {
			const node = view.state.doc.child(index);
			return paragraphFromAttrs(
				node.attrs,
				String(node.attrs.id),
				Array.from({ length: node.childCount }, (_, run) => inlineNodeRun(node.child(run))!),
			);
		});
		const native = await loadDocx(
			new Uint8Array(
				await readFile(
					new URL(`../__fixtures__/paragraph-boundary-tracking/${action}.docx`, import.meta.url),
				),
			),
		);
		const project = (paragraphs: typeof source.model.blocks) =>
			paragraphs.map((block) => {
				if (block.type !== 'paragraph') throw new Error('Expected paragraph');
				return {
					text: block.runs.map((run) => run.text).join(''),
					kind: block.markRevision?.kind,
					align: block.align,
					keepNext: block.keepNext,
					spacing: block.spacingAfterTwips,
					bold: block.runs.every((run) => run.bold),
					formatRevision: block.formatRevision?.kind,
				};
			});
		expect(project(blocks)).toEqual(project(native.model.blocks));
		expect(
			project((await loadDocx(await source.save({ ...source.model, blocks }))).model.blocks),
		).toEqual(project(native.model.blocks));
	},
);
