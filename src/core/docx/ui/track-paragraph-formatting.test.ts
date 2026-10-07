import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { history, undo, redo } from 'prosemirror-history';
import { loadDocx, saveDocx, rejectAllRevisions, listRevisions } from '../index';
import type { Paragraph } from '../model';
import { PARAGRAPH_FORMAT_KEYS } from '../restore-paragraph-format';
import { propertiesSignature } from '../revision-properties';
import { first, parseXml, WORD_NS } from '../xml';
import { paragraphAttrs, paragraphFromAttrs } from './paragraph-attributes';
import { marksForRun } from './run-marks';
import { markSpecs } from './schema-marks';
import {
	REMOTE_TRANSACTION_META,
	trackChangesPlugin,
	trackChangesPluginKey,
} from './track-changes-mode';

const empty: Paragraph = { type: 'paragraph', id: 'empty', runs: [] };
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+', attrs: { trackFormatting: { default: true } } },
		paragraph: {
			content: 'text*',
			attrs: Object.fromEntries(
				Object.entries(paragraphAttrs(empty)).map(([key, value]) => [key, { default: value }]),
			),
		},
		text: {},
	},
	marks: markSpecs,
});
function state(paragraphs: Paragraph[] = [empty], enabled = true) {
	return EditorState.create({
		doc: schema.node(
			'doc',
			null,
			paragraphs.map((paragraph) =>
				schema.node(
					'paragraph',
					paragraphAttrs(paragraph),
					paragraph.runs.map((run) => schema.text(run.text, marksForRun(run, schema))),
				),
			),
		),
		plugins: [
			history(),
			trackChangesPlugin(
				() => 'Ada',
				() => enabled,
			),
		],
	});
}
function modelParagraph(editor: EditorState, original: Paragraph): Paragraph {
	return paragraphFromAttrs(editor.doc.firstChild!.attrs, original.id, original.runs);
}
const fixture = async (name: string) =>
	loadDocx(
		new Uint8Array(
			await readFile(
				new URL(`../__fixtures__/review-paragraph-formatting/${name}.docx`, import.meta.url),
			),
		),
	);
const paragraphOf = (model: Awaited<ReturnType<typeof loadDocx>>['model']) => {
	const paragraph = model.blocks[0]!;
	if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
	return paragraph;
};
const formatting = (paragraph: Paragraph) =>
	Object.fromEntries(PARAGRAPH_FORMAT_KEYS.map((key) => [key, paragraph[key]]));

describe('shared paragraph formatting recording', () => {
	it('retains imported native history through later edits and clears it on full reversion', async () => {
		const before = paragraphOf((await fixture('alignment-before')).model);
		const tracked = paragraphOf((await fixture('alignment-tracked')).model);
		let editor = state([tracked]);
		editor = editor.apply(editor.tr.setNodeAttribute(0, 'keepNext', true));
		expect(editor.doc.firstChild!.attrs.formatRevision).toEqual(tracked.formatRevision);
		editor = editor.apply(
			editor.tr
				.setNodeAttribute(0, 'keepNext', before.keepNext ?? null)
				.setNodeAttribute(0, 'align', before.align ?? null),
		);
		expect(editor.doc.firstChild!.attrs.formatRevision).toBeNull();
	});
	it('does not introduce history for a formatting no-op', () => {
		const editor = state([{ ...empty, align: 'center' }]);
		const after = editor.apply(editor.tr.setNodeAttribute(0, 'align', 'center'));
		expect(after.doc.firstChild!.attrs.formatRevision).toBeNull();
	});
	it.each(['alignment', 'spacing', 'indent', 'multiple'])(
		'matches native %s actions and rejects both export paths',
		async (name) => {
			const loaded = await fixture(`${name}-before`);
			const before = paragraphOf(loaded.model);
			const native = paragraphOf((await fixture(`${name}-tracked`)).model);
			let editor = state([before]);
			const attrs = { ...paragraphAttrs(native) };
			for (const key of [
				'id',
				'formatRevision',
				'markRevision',
				'sourceParagraphPropertiesXml',
				'restoredParagraphPropertiesXml',
			])
				attrs[key] = editor.doc.firstChild!.attrs[key];
			editor = editor.apply(editor.tr.setNodeMarkup(0, undefined, attrs));
			const tracked = modelParagraph(editor, before);
			expect(formatting(tracked)).toEqual(formatting(native));
			expect(tracked.formatRevision).toMatchObject({
				kind: 'paragraphChange',
				author: 'Ada',
				dateUtc: tracked.formatRevision?.date,
			});
			const prior = parseXml(
				tracked.formatRevision!.previousParagraphPropertiesXml!,
			).documentElement;
			expect(first(prior, 'rPr')).toBeUndefined();
			expect(first(prior, 'sectPr')).toBeUndefined();
			loaded.model.blocks[0] = tracked;
			for (const bytes of [await loaded.save(loaded.model), await saveDocx(loaded.model)]) {
				const exported = (await loadDocx(bytes)).model;
				expect(listRevisions(exported).map((revision) => revision.kind)).toEqual([
					'paragraphChange',
				]);
				const rejected = paragraphOf(rejectAllRevisions(exported));
				expect(formatting(rejected)).toEqual(formatting(before));
				expect(
					propertiesSignature(
						first(parseXml(rejected.sourceParagraphPropertiesXml!).documentElement, 'rPr')!,
					),
				).toBe(
					propertiesSignature(
						first(parseXml(before.sourceParagraphPropertiesXml!).documentElement, 'rPr')!,
					),
				);
			}
		},
	);
	it('retains the earliest snapshot and removes history when formatting returns to it', () => {
		let editor = state();
		editor = editor.apply(editor.tr.setNodeAttribute(0, 'align', 'center'));
		const revision = editor.doc.firstChild!.attrs.formatRevision;
		editor = editor.apply(editor.tr.setNodeAttribute(0, 'keepNext', true));
		expect(editor.doc.firstChild!.attrs.formatRevision).toEqual(revision);
		editor = editor.apply(
			editor.tr.setNodeAttribute(0, 'align', null).setNodeAttribute(0, 'keepNext', null),
		);
		expect(editor.doc.firstChild!.attrs.formatRevision).toBeNull();
	});
	it('includes recording in atomic undo and redo for an empty paragraph', () => {
		let editor = state();
		editor = editor.apply(editor.tr.setNodeAttribute(0, 'align', 'center'));
		const tracked = editor.doc;
		expect(
			undo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.firstChild!.attrs.align).toBeNull();
		expect(editor.doc.firstChild!.attrs.formatRevision).toBeNull();
		expect(
			redo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.eq(tracked)).toBe(true);
	});
	it('keeps separate prior snapshots for paragraphs changed by one action', () => {
		let editor = state([empty, { ...empty, id: 'second', align: 'right' }]);
		editor = editor.apply(
			editor.tr.setNodeAttribute(0, 'align', 'center').setNodeAttribute(2, 'align', 'center'),
		);
		const firstRevision = editor.doc.firstChild!.attrs.formatRevision;
		const secondRevision = editor.doc.lastChild!.attrs.formatRevision;
		expect(firstRevision.id).toBe(secondRevision.id);
		expect(firstRevision.previousParagraphPropertiesXml).not.toBe(
			secondRevision.previousParagraphPropertiesXml,
		);
	});
	it.each(['disabled', 'preference', 'remote', 'review', 'identity', 'mixed'])(
		'skips %s transactions',
		(kind) => {
			let editor = state([{ ...empty, runs: [{ text: 'Text' }] }], kind !== 'disabled');
			if (kind === 'preference')
				editor = editor.apply(editor.tr.setDocAttribute('trackFormatting', false));
			let tr = editor.tr.setNodeAttribute(0, 'align', 'center');
			if (kind === 'remote') tr.setMeta(REMOTE_TRANSACTION_META, true);
			if (kind === 'review') tr.setMeta(trackChangesPluginKey, { tracked: true });
			if (kind === 'identity') tr.setNodeAttribute(0, 'id', 'other');
			if (kind === 'mixed') tr.insertText('!', 1);
			editor = editor.apply(tr);
			expect(editor.doc.firstChild!.attrs.formatRevision).toBeNull();
		},
	);
	it('preserves opaque properties and excludes independent paragraph-mark and section history', () => {
		const source = `<w:pPr xmlns:w="${WORD_NS}" xmlns:x="urn:extra"><w:spacing w:after="0" x:extra="keep"/><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:sectPr><w:cols w:num="2"/></w:sectPr></w:pPr>`;
		let editor = state([
			{
				...empty,
				spacingAfterTwips: 0 as NonNullable<Paragraph['spacingAfterTwips']>,
				sourceParagraphPropertiesXml: source,
			},
		]);
		editor = editor.apply(editor.tr.setNodeAttribute(0, 'align', 'center'));
		const prior = editor.doc.firstChild!.attrs.formatRevision.previousParagraphPropertiesXml;
		expect(prior).toContain('x:extra="keep"');
		expect(prior).not.toContain('rFonts');
		expect(prior).not.toContain('sectPr');
		expect(editor.doc.firstChild!.attrs.sourceParagraphPropertiesXml).toBe(source);
	});
});
