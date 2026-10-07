import { expect, it } from 'vitest';
import { Fragment, Schema, Slice } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import type { TextRun } from '../model';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { markSpecs } from './schema-marks';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';
import { fieldGuardPlugin, FIELD_RESULT_REPAIR_META } from './field-guard';
import { collectRevisionRanges } from './review-commands';
import { REMOTE_TRANSACTION_META, trackChangesPlugin } from './track-changes-mode';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+', attrs: { trackFormatting: { default: true } } },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});

function field(locked: boolean | undefined, result: TextRun[] = []): TextRun[] {
	return [
		{ text: '', fieldChar: 'begin', ...(locked === undefined ? {} : { fieldFlags: { locked } }) },
		{ text: '', fieldCode: 'REF Target' },
		{ text: '', fieldChar: 'separate' },
		...result,
		{ text: '', fieldChar: 'end' },
	];
}
function editor(runs: TextRun[], tracked = false, formatting = true) {
	return EditorState.create({
		doc: schema.node(
			'doc',
			{ trackFormatting: formatting },
			schema.node(
				'paragraph',
				null,
				runs.flatMap((run) => runToInlineNodes(run, schema)),
			),
		),
		plugins: [
			fieldGuardPlugin(),
			trackChangesPlugin(
				() => 'Ada',
				() => tracked,
			),
		],
	});
}
function results(state: EditorState) {
	const runs: TextRun[] = [];
	state.doc.descendants((node) => {
		if (node.isText) runs.push(inlineNodeRun(node)!);
	});
	return runs;
}

for (const locked of [true, false, undefined])
	it(`projects begin lock ${locked} onto existing results while preserving opaque formatting`, () => {
		const state = editor(
			field(locked, [
				{
					text: 'Old',
					field: { instr: 'REF Target' },
					fieldFlags: { locked: locked !== true, dirty: false },
					bold: true,
					sourceRunPropertiesXml: '<w:rPr><w:custom/></w:rPr>',
				},
			]),
		);
		const applied = state.applyTransaction(state.tr.insertText('!', 5));
		expect(applied.transactions.some((tr) => tr.getMeta(FIELD_RESULT_REPAIR_META))).toBe(true);
		for (const run of results(applied.state)) {
			expect(run.fieldFlags?.locked).toBe(locked);
			expect(run.bold).toBe(true);
			expect(run.fieldFlags?.dirty).toBe(false);
			expect(run.sourceRunPropertiesXml).toBe('<w:rPr><w:custom/></w:rPr>');
		}
	});

it('deleting every cached character and retyping restores the authoritative lock', () => {
	let state = editor(
		field(true, [{ text: 'Old', field: { instr: 'REF Target' }, fieldFlags: { locked: true } }]),
	);
	state = state.applyTransaction(state.tr.delete(4, 7)).state;
	state = state.applyTransaction(state.tr.insertText('New', 4)).state;
	expect(results(state)).toMatchObject([
		{ text: 'New', field: { instr: 'REF Target' }, fieldFlags: { locked: true } },
	]);
});

it('removes stale lock-only metadata when the begin marker has no lock attribute', () => {
	const state = editor(
		field(undefined, [
			{ text: 'Old', field: { instr: 'REF Target' }, fieldFlags: { locked: true } },
		]),
	);
	const next = state.applyTransaction(state.tr.insertText('!', 5)).state;
	expect(results(next)).toEqual([{ text: 'O!ld', field: { instr: 'REF Target' } }]);
	expect(next.doc.nodeAt(4)!.marks.some((mark) => mark.type.name === 'runProperties')).toBe(false);
});

it('tracks whole cached result replacement with the begin lock on both revision sides', () => {
	const state = editor(
		field(true, [{ text: 'Old', field: { instr: 'REF Target' }, fieldFlags: { locked: true } }]),
		true,
	);
	const next = state.applyTransaction(state.tr.insertText('New', 4, 7)).state;
	expect(collectRevisionRanges(next.doc).map((range) => range.kind)).toEqual(['delete', 'insert']);
	expect(
		results(next).map((run) => [run.text, run.fieldFlags?.locked, run.revision?.kind]),
	).toEqual([
		['Old', true, 'delete'],
		['New', true, 'insert'],
	]);
});

it('result-only paste preserves the pasted formatting and replaces its foreign lock', () => {
	const state = editor(field(true));
	const slice = new Slice(
		Fragment.fromArray(
			runToInlineNodes(
				{ text: 'Paste', italic: true, fieldFlags: { locked: false, dirty: true } },
				schema,
			),
		),
		0,
		0,
	);
	const next = state.applyTransaction(
		state.tr.replace(4, 4, slice).setMeta('uiEvent', 'paste'),
	).state;
	expect(results(next)).toMatchObject([
		{ text: 'Paste', italic: true, fieldFlags: { locked: true, dirty: true } },
	]);
});

it('nested field locks are independent and the outer result resumes after the inner end', () => {
	const state = editor([
		...field(true, [
			{ text: 'Outer' },
			...field(false, [{ text: 'Inner', fieldFlags: { locked: true } }]),
			{ text: 'Tail' },
		]),
	]);
	const next = state.applyTransaction(state.tr.insertText('!', 4)).state;
	expect(results(next).map((run) => [run.text, run.fieldFlags?.locked])).toEqual([
		['!Outer', true],
		['Inner', false],
		['Tail', true],
	]);
});

for (const formatting of [true, false])
	it(`tracks empty result typing once with formatting tracking ${formatting}`, () => {
		const state = editor(field(true), true, formatting);
		const next = state.applyTransaction(state.tr.insertText('New', 4)).state;
		expect(collectRevisionRanges(next.doc).map((range) => range.kind)).toEqual(['insert']);
		expect(results(next)).toMatchObject([
			{ text: 'New', fieldFlags: { locked: true }, revision: { kind: 'insert', author: 'Ada' } },
		]);
		expect(results(next).every((run) => !run.formatRevision)).toBe(true);
	});

it('repairs remotely inserted result metadata without locally tracking it', () => {
	const state = editor(field(true), true);
	const next = state.applyTransaction(
		state.tr.insertText('Remote', 4).setMeta(REMOTE_TRANSACTION_META, true),
	).state;
	expect(results(next)).toMatchObject([{ text: 'Remote', fieldFlags: { locked: true } }]);
	expect(collectRevisionRanges(next.doc)).toEqual([]);
});

it('begin metadata changes repair existing marked results without a formatting revision', () => {
	const state = editor(
		field(false, [
			{ text: 'Old', field: { instr: 'REF Target' }, fieldFlags: { locked: false }, bold: true },
		]),
		true,
	);
	const begin = state.doc.nodeAt(1)!;
	const next = state.applyTransaction(
		state.tr.setNodeAttribute(
			1,
			'format',
			JSON.stringify({ ...JSON.parse(begin.attrs.format), fieldFlags: { locked: true } }),
		),
	).state;
	expect(results(next)).toMatchObject([{ text: 'Old', bold: true, fieldFlags: { locked: true } }]);
	expect(collectRevisionRanges(next.doc)).toEqual([]);
});
