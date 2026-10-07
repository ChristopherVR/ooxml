import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { loadDocx, saveDocx, rejectAllRevisions, listRevisions } from '../index';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { history, undo, redo } from 'prosemirror-history';
import type { TextRun } from '../model';
import { first, parseXml, WORD_NS } from '../xml';
import { marksForRun } from './run-marks';
import { applyMarkFormatting } from './run-mark-properties';
import { markSpecs } from './schema-marks';
import { formattingRevision, resolveFormattingRange } from './review-formatting';
import {
	REMOTE_TRANSACTION_META,
	trackChangesPlugin,
	trackChangesPluginKey,
} from './track-changes-mode';

const schema = new Schema({
	nodes: {
		doc: {
			content: 'paragraph+',
			attrs: { trackFormatting: { default: true }, trackMoves: { default: true } },
		},
		paragraph: { content: 'text*' },
		text: {},
	},
	marks: markSpecs,
});
function state(
	runs: TextRun[] = [{ text: 'Format me', bold: false, italic: false }],
	enabled = true,
) {
	return EditorState.create({
		doc: schema.node(
			'doc',
			null,
			schema.node(
				'paragraph',
				null,
				runs.map((run) => schema.text(run.text, marksForRun(run, schema))),
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
function runAt(editor: EditorState, pos = 1): TextRun {
	const node = editor.doc.nodeAt(pos)!;
	const run: TextRun = { text: node.text! };
	applyMarkFormatting(run, node);
	return run;
}

describe('shared formatting revision recording', () => {
	it('honors native disabled formatting tracking while continuing to record insertions', async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(
					new URL('../__fixtures__/review-preferences/preferences.docx', import.meta.url),
				),
			),
		);
		const paragraph = loaded.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		let editor = state(paragraph.runs);
		editor = editor.apply(
			editor.tr.setDocAttribute('trackFormatting', loaded.model.trackFormatting),
		);
		editor = editor.apply(editor.tr.addMark(1, 11, schema.marks.bold!.create()));
		expect(runAt(editor).bold).toBe(true);
		expect(runAt(editor).formatRevision).toBeUndefined();
		editor = editor.apply(editor.tr.insertText('!', 16));
		expect(runAt(editor, 16).revision?.kind).toBe('insert');
	});
	it('records a drag as insertion/deletion when move recording is disabled', () => {
		let editor = state();
		editor = editor.apply(editor.tr.setDocAttribute('trackMoves', false));
		const tr = editor.tr.delete(1, 2);
		tr.insertText('F', tr.mapping.map(10));
		editor = editor.apply(tr.setMeta('uiEvent', 'drop'));
		const source = runAt(editor).revision;
		const target = runAt(editor, 10).revision;
		expect(source?.kind).toBe('delete');
		expect(target?.kind).toBe('insert');
		expect(source?.move).toBeUndefined();
		expect(target?.move).toBeUndefined();
	});
	it.each(['bold', 'bold-italic', 'bold-off', 'own-insertion'])(
		'matches native Word %s recording and exports rejectable snapshots',
		async (action) => {
			const fixture = (name: string) =>
				new URL(`../__fixtures__/formatting-actions/${name}.docx`, import.meta.url);
			const loaded = await loadDocx(new Uint8Array(await readFile(fixture('baseline'))));
			const paragraph = loaded.model.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			const original = structuredClone(paragraph.runs);
			let editor = state(original);
			if (action === 'own-insertion') editor = editor.apply(editor.tr.insertText('!', 10));
			const from = action === 'own-insertion' ? 10 : 1;
			const to = action === 'own-insertion' ? 11 : 7;
			editor = editor.apply(editor.tr.addMark(from, to, schema.marks.bold!.create()));
			if (action === 'bold-italic')
				editor = editor.apply(editor.tr.addMark(from, to, schema.marks.italic!.create()));
			if (action === 'bold-off')
				editor = editor.apply(editor.tr.removeMark(from, to, schema.marks.bold!));
			const runs: TextRun[] = [];
			editor.doc.descendants((node) => {
				if (!node.isText) return;
				const run: TextRun = { text: node.text! };
				applyMarkFormatting(run, node);
				runs.push(run);
			});
			paragraph.runs = runs;
			const native = await loadDocx(new Uint8Array(await readFile(fixture(action))));
			const kinds = (model: typeof native.model) =>
				listRevisions(model)
					.map((revision) => revision.kind)
					.sort();
			expect(kinds(loaded.model)).toEqual(kinds(native.model));
			for (const bytes of [await loaded.save(loaded.model), await saveDocx(loaded.model)]) {
				const exported = await loadDocx(bytes);
				expect(kinds(exported.model)).toEqual(kinds(native.model));
				const rejected = rejectAllRevisions(exported.model).blocks[0]!;
				if (rejected.type !== 'paragraph') throw new Error('Expected paragraph');
				expect(rejected.runs.map((run) => run.text).join('')).toBe('Format me');
				for (const run of rejected.runs) {
					expect(Boolean(run.bold)).toBe(false);
					expect(Boolean(run.italic)).toBe(false);
					expect(run.fontFamily).toBe(original[0]!.fontFamily);
				}
			}
		},
	);
	it('records complete prior formatting and rejects only the selected formatting', () => {
		let editor = state([
			{ text: 'Format me', bold: false, fontFamilyComplexScript: 'Amiri', commentIds: ['comment'] },
		]);
		editor = editor.apply(editor.tr.addMark(1, 7, schema.marks.bold!.create()));
		const run = runAt(editor);
		expect(run.bold).toBe(true);
		expect(run.formatRevision).toMatchObject({ kind: 'formatChange', author: 'Ada' });
		expect(run.formatRevision!.dateUtc).toBe(run.formatRevision!.date);
		const tr = editor.tr;
		resolveFormattingRange(tr, 1, 7, 'reject');
		editor = editor.apply(tr.setMeta(trackChangesPluginKey, { tracked: true }));
		expect(runAt(editor)).toMatchObject({
			bold: false,
			fontFamilyComplexScript: 'Amiri',
			commentIds: ['comment'],
		});
		expect(formattingRevision(editor.doc.nodeAt(1)!)).toBeUndefined();
		expect(editor.doc.textContent).toBe('Format me');
	});
	it('keeps the earliest snapshot for successive changes and removes reverted history', () => {
		let editor = state();
		editor = editor.apply(editor.tr.addMark(1, 7, schema.marks.bold!.create()));
		const revision = runAt(editor).formatRevision;
		editor = editor.apply(editor.tr.addMark(1, 7, schema.marks.italic!.create()));
		expect(runAt(editor).formatRevision).toEqual(revision);
		editor = editor.apply(editor.tr.removeMark(1, 7, schema.marks.bold!));
		expect(runAt(editor).formatRevision).toEqual(revision);
		editor = editor.apply(editor.tr.removeMark(1, 7, schema.marks.italic!));
		expect(runAt(editor).formatRevision).toBeUndefined();
	});
	it('retains an independent snapshot for each original run that merges after formatting', () => {
		let editor = state([
			{ text: 'A', bold: false },
			{ text: 'B', bold: true },
		]);
		editor = editor.apply(editor.tr.addMark(1, 3, schema.marks.bold!.create()));
		expect(runAt(editor).formatRevision).toBeDefined();
		expect(runAt(editor, 2).formatRevision).toBeUndefined();
		const tr = editor.tr;
		resolveFormattingRange(tr, 1, 3, 'reject');
		editor = editor.apply(tr.setMeta(trackChangesPluginKey, { tracked: true }));
		expect(runAt(editor).bold).toBe(false);
		expect(runAt(editor, 2).bold).toBe(true);
	});
	it('records formatting alongside the author’s pending insertion', () => {
		let editor = state();
		editor = editor.apply(editor.tr.insertText('!', 10));
		const textRevision = runAt(editor, 10).revision;
		editor = editor.apply(editor.tr.addMark(10, 11, schema.marks.bold!.create()));
		expect(runAt(editor, 10).revision).toEqual(textRevision);
		expect(runAt(editor, 10).formatRevision?.kind).toBe('formatChange');
	});
	it('tracks UI toggle transactions that replace explicit-off properties alongside bold', () => {
		let editor = state([{ text: 'Text', bold: false, fontFamilyComplexScript: 'Amiri' }]);
		const tr = editor.tr.removeMark(1, 5, schema.marks.runProperties!);
		tr.addMark(
			1,
			5,
			schema.marks.runProperties!.create({ props: { fontFamilyComplexScript: 'Amiri' } }),
		);
		tr.addMark(1, 5, schema.marks.bold!.create());
		editor = editor.apply(tr);
		expect(runAt(editor).formatRevision?.kind).toBe('formatChange');
		const reject = editor.tr;
		resolveFormattingRange(reject, 1, 5, 'reject');
		editor = editor.apply(reject.setMeta(trackChangesPluginKey, { tracked: true }));
		expect(runAt(editor)).toMatchObject({ bold: false, fontFamilyComplexScript: 'Amiri' });
	});
	it('undoes and redoes the formatting and its history atomically', () => {
		let editor = state();
		const original = editor.doc;
		editor = editor.apply(editor.tr.addMark(1, 7, schema.marks.bold!.create()));
		const changed = editor.doc;
		expect(
			undo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.eq(original)).toBe(true);
		expect(
			redo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.eq(changed)).toBe(true);
	});
	it('preserves opaque properties in the prior snapshot and current marks', () => {
		const xml = `<w:rPr xmlns:w="${WORD_NS}"><w:outline/><w:shadow/></w:rPr>`;
		let editor = state([{ text: 'Text', sourceRunPropertiesXml: xml }]);
		editor = editor.apply(editor.tr.addMark(1, 5, schema.marks.bold!.create()));
		const run = runAt(editor);
		expect(run.sourceRunPropertiesXml).toBe(xml);
		const props = parseXml(run.formatRevision!.previousRunPropertiesXml!).documentElement;
		expect(first(props, 'outline')).toBeDefined();
		expect(first(props, 'shadow')).toBeDefined();
	});
	it.each(['disabled', 'remote', 'review', 'link', 'noop'])(
		'does not record %s changes',
		(mode) => {
			let editor = state(undefined, mode !== 'disabled');
			const mark =
				mode === 'link'
					? schema.marks.link!.create({ href: 'https://example.com' })
					: schema.marks.bold!.create();
			let tr = editor.tr.addMark(1, 7, mark);
			if (mode === 'remote') tr = tr.setMeta(REMOTE_TRANSACTION_META, true);
			if (mode === 'review') tr = tr.setMeta(trackChangesPluginKey, { tracked: true });
			if (mode === 'noop') tr = tr.removeMark(1, 7, schema.marks.bold!);
			editor = editor.apply(tr);
			expect(runAt(editor).formatRevision).toBeUndefined();
		},
	);
});
