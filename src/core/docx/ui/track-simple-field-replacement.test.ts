import { expect, it } from 'vitest';
import { Fragment, Schema, Slice } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { ReplaceStep } from 'prosemirror-transform';
import JSZip from 'jszip';
import { loadDocx, type Paragraph, type TextRun } from '../index';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { markSpecs } from './schema-marks';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';
import { fieldGuardPlugin } from './field-guard';
import { replaceSimpleFieldResult, simpleFieldPasteSlice } from './simple-field-input';
import { trackChangesPlugin } from './track-changes-mode';
import { acceptAllChanges, rejectAllChanges, collectRevisionRanges } from './review-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});

async function saveResults(runs: TextRun[]) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>L</w:t></w:r><w:fldSimple w:instr=" QUOTE ABCD "><w:r><w:rPr><w:b/></w:rPr><w:t>ABCD</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" QUOTE EF "><w:r><w:t>EF</w:t></w:r></w:fldSimple><w:r><w:t>R</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const saved = await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] });
	const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
	expect(xml.match(/QUOTE ABCD/g)).toHaveLength(1);
	expect(xml.match(/QUOTE EF/g)).toHaveLength(1);
	return (await loadDocx(saved)).model.blocks[0] as Paragraph;
}

for (const input of ['type', 'paste'] as const)
	for (const whole of [false, true])
		for (const ownInsertion of [false, true])
			for (const action of ['accept', 'reject'] as const)
				it(`${action} tracked ${input} ${whole ? 'whole' : 'partial'} simple result (${ownInsertion ? 'own insertion' : 'imported'})`, async () => {
					const first = runToInlineNodes(
						{
							text: 'ABCD',
							bold: true,
							field: { instr: 'QUOTE ABCD', simple: true },
							fieldInstanceId: 'first',
							...(ownInsertion
								? { revision: { kind: 'insert' as const, author: 'Ada', id: 'prior' } }
								: {}),
						},
						schema,
					);
					const adjacent = runToInlineNodes(
						{
							text: 'EF',
							field: { instr: 'QUOTE EF', simple: true },
							fieldInstanceId: 'second',
						},
						schema,
					);
					const doc = schema.node(
						'doc',
						null,
						schema.node('paragraph', null, [
							schema.text('L'),
							...first,
							...adjacent,
							schema.text('R'),
						]),
					);
					const from = whole ? 2 : 3;
					const to = whole ? 6 : 5;
					let state = EditorState.create({
						doc,
						selection: TextSelection.create(doc, from, to),
						plugins: [
							fieldGuardPlugin(),
							trackChangesPlugin(
								() => 'Ada',
								() => true,
							),
						],
					});
					let tr: Transaction;
					if (input === 'type') {
						tr = replaceSimpleFieldResult(state, from, to, 'X')!;
					} else {
						const slice = simpleFieldPasteSlice(
							new Slice(Fragment.from(schema.text('X', [schema.marks.italic!.create()])), 0, 0),
							state,
						);
						tr = state.tr.replaceSelection(slice).setMeta('uiEvent', 'paste');
					}
					expect(tr.steps.every((step) => step instanceof ReplaceStep)).toBe(true);
					state = state.applyTransaction(tr).state;
					const inserted = collectRevisionRanges(state.doc).filter(
						(range) => range.kind === 'insert',
					);
					expect(
						inserted.some((range) => state.doc.textBetween(range.from, range.to) === 'X'),
					).toBe(true);
					const deleted = collectRevisionRanges(state.doc).filter(
						(range) => range.kind === 'delete',
					);
					expect(deleted).toHaveLength(ownInsertion ? 0 : 1);
					if (!ownInsertion)
						expect(state.doc.textBetween(deleted[0]!.from, deleted[0]!.to)).toBe(
							whole ? 'ABCD' : 'BC',
						);
					const view = {
						get state() {
							return state;
						},
						editable: true,
						dispatch(transaction: Transaction) {
							state = state.applyTransaction(transaction).state;
						},
					} as EditorView;
					(action === 'accept' ? acceptAllChanges : rejectAllChanges)(view);
					expect(state.doc.textContent).toBe(
						action === 'accept'
							? whole
								? 'LXEFR'
								: 'LAXDEFR'
							: ownInsertion
								? 'LEFR'
								: 'LABCDEFR',
					);
					const runs = Array.from({ length: state.doc.firstChild!.childCount }, (_, index) =>
						inlineNodeRun(state.doc.firstChild!.child(index))!,
					);
					expect(collectRevisionRanges(state.doc)).toEqual([]);
					expect(runs.filter((run) => run.fieldInstanceId === 'second')).toEqual([
						{ text: 'EF', field: { instr: 'QUOTE EF', simple: true }, fieldInstanceId: 'second' },
					]);
					const firstResults = runs.filter((run) => run.fieldInstanceId === 'first');
					if (action === 'reject' && ownInsertion) {
						expect(runs.filter((run) => run.fieldCode === 'QUOTE ABCD')).toHaveLength(1);
						expect(runs.filter((run) => run.fieldChar === 'begin')).toHaveLength(1);
					} else {
						expect(firstResults.length).toBeGreaterThan(0);
						expect(firstResults.every((run) => run.field?.instr === 'QUOTE ABCD')).toBe(true);
						if (action === 'accept' && input === 'paste')
							expect(firstResults.find((run) => run.text.includes('X'))?.italic).toBe(true);
					}
					const reloaded = await saveResults(runs);
					expect(reloaded.runs.map((run) => run.text).join('')).toBe(state.doc.textContent);
					expect(reloaded.runs.some((run) => run.revision)).toBe(false);
				});

for (const whole of [false, true])
	for (const action of ['accept', 'reject'] as const)
		it(`${action} replacement removes only the author's replaced pending portion (${whole ? 'whole' : 'partial'})`, async () => {
			const run = (text: string, pending = false) =>
				runToInlineNodes(
					{
						text,
						field: { instr: 'QUOTE ABCD', simple: true },
						fieldInstanceId: 'first',
						...(pending
							? { revision: { kind: 'insert' as const, author: 'Ada', id: 'prior' } }
							: {}),
					},
					schema,
				);
			const doc = schema.node(
				'doc',
				null,
				schema.node('paragraph', null, [
					schema.text('L'),
					...run('A'),
					...run('BC', true),
					...run('D'),
					schema.text('R'),
				]),
			);
			const from = whole ? 2 : 3;
			const to = whole ? 6 : 5;
			let state = EditorState.create({
				doc,
				selection: TextSelection.create(doc, from, to),
				plugins: [
					fieldGuardPlugin(),
					trackChangesPlugin(
						() => 'Ada',
						() => true,
					),
				],
			});
			state = state.applyTransaction(replaceSimpleFieldResult(state, from, to, 'X')!).state;
			const ranges = collectRevisionRanges(state.doc);
			expect(ranges.some((range) => range.id === 'prior')).toBe(false);
			expect(
				ranges
					.filter((range) => range.kind === 'delete')
					.map((range) => state.doc.textBetween(range.from, range.to))
					.join(''),
			).toBe(whole ? 'AD' : '');
			const view = {
				get state() {
					return state;
				},
				editable: true,
				dispatch(tr: Transaction) {
					state = state.applyTransaction(tr).state;
				},
			} as EditorView;
			(action === 'accept' ? acceptAllChanges : rejectAllChanges)(view);
			expect(state.doc.textContent).toBe(action === 'reject' ? 'LADR' : whole ? 'LXR' : 'LAXDR');
			expect(collectRevisionRanges(state.doc)).toEqual([]);
			const runs = Array.from({ length: state.doc.firstChild!.childCount }, (_, index) =>
				inlineNodeRun(state.doc.firstChild!.child(index))!,
			);
			expect(
				runs
					.filter((item) => item.field)
					.every((item) => item.field?.instr === 'QUOTE ABCD' && item.fieldInstanceId === 'first'),
			).toBe(true);
		});

it('tracks text inserted with explicit stored formatting without separate metadata mark steps', () => {
	const nodes = runToInlineNodes(
		{ text: 'AB', field: { instr: 'QUOTE AB', simple: true }, fieldInstanceId: 'first' },
		schema,
	);
	const doc = schema.node('doc', null, schema.node('paragraph', null, nodes));
	let state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, 2),
		plugins: [
			fieldGuardPlugin(),
			trackChangesPlugin(
				() => 'Ada',
				() => true,
			),
		],
	});
	state = state.apply(state.tr.setStoredMarks([schema.marks.italic!.create()]));
	const tr = replaceSimpleFieldResult(state, 2, 2, 'X')!;
	expect(tr.steps).toHaveLength(1);
	expect(tr.steps[0]).toBeInstanceOf(ReplaceStep);
	state = state.applyTransaction(tr).state;
	expect(collectRevisionRanges(state.doc)).toHaveLength(1);
	expect(inlineNodeRun(state.doc.nodeAt(2)!)).toMatchObject({
		text: 'X',
		italic: true,
		field: { instr: 'QUOTE AB', simple: true },
		fieldInstanceId: 'first',
		revision: { kind: 'insert', author: 'Ada' },
	});
});
