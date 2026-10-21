import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import type { ContentControl, DocumentProtection, FormField, TextRun } from '../model';
import { setFormCheckbox } from '../form-fields';
import { contentControlPlugin, violatesControlLock } from './content-control-guard';
import { markSpecs } from './schema-marks';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { appendInlineNode, runToInlineNodes } from './run-adapter';
import {
	formFieldLocked,
	formFieldNodeDisplay,
	formFieldOfNode,
	formTextResults,
	selectFormDropDownAt,
	toggleFormCheckboxAt,
} from './form-field-commands';
import { formFieldGuardPlugin } from './form-field-guard';
import { sameRuns } from './run-compare';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: {
			content: 'inline*',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['p', 0],
		},
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});

const field = (formField: FormField, result?: string): TextRun[] => {
	const instr =
		formField.kind === 'text'
			? 'FORMTEXT'
			: formField.kind === 'checkbox'
				? 'FORMCHECKBOX'
				: 'FORMDROPDOWN';
	return [
		{ text: '', fieldChar: 'begin', formField },
		{ text: '', fieldCode: ` ${instr} ` },
		...(result === undefined
			? []
			: [{ text: '', fieldChar: 'separate' } as TextRun, { text: result, field: { instr } }]),
		{ text: '', fieldChar: 'end' },
	];
};
const name: FormField = { kind: 'text', name: 'Name', textInput: { maxLength: 6 } };
const approved: FormField = { kind: 'checkbox', name: 'Approved', checkBox: { default: false } };
const status: FormField = {
	kind: 'dropDown',
	name: 'Status',
	dropDown: { entries: ['New', 'Revised'], result: 0 },
};

function stateOf(protection?: DocumentProtection) {
	const runs: TextRun[] = [
		{ text: 'Name: ' },
		...field(name, 'Ada'),
		{ text: ' Approved ' },
		...field(approved),
		{ text: ' Status ' },
		...field(status),
	];
	const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
	const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
	return EditorState.create({ schema, doc, plugins: [formFieldGuardPlugin(() => protection)] });
}

function markerPos(state: EditorState, kind: FormField['kind']): number {
	let found = -1;
	state.doc.descendants((node, pos) => {
		if (formFieldOfNode(node)?.kind === kind) found = pos;
	});
	return found;
}

const runsOf = (state: EditorState): TextRun[] => {
	const runs: TextRun[] = [];
	state.doc.firstChild!.forEach((child) => appendInlineNode(runs, child));
	return runs;
};

describe('legacy form fields in the editor', () => {
	it('draws a check box and a drop-down from their field data', () => {
		const state = stateOf();
		expect(formFieldNodeDisplay(state.doc.nodeAt(markerPos(state, 'checkbox'))!)).toBe('☐');
		expect(formFieldNodeDisplay(state.doc.nodeAt(markerPos(state, 'dropDown'))!)).toBe('New');
		expect(formFieldNodeDisplay(state.doc.nodeAt(markerPos(state, 'text'))!)).toBeUndefined();
	});

	it('toggles a check box and picks a drop-down entry, carried back to the run model', () => {
		let state = stateOf();
		state = state.apply(toggleFormCheckboxAt(state, markerPos(state, 'checkbox'))!);
		state = state.apply(selectFormDropDownAt(state, markerPos(state, 'dropDown'), 1)!);
		expect(selectFormDropDownAt(state, markerPos(state, 'dropDown'), 5)).toBeUndefined();
		expect(toggleFormCheckboxAt(state, markerPos(state, 'dropDown'))).toBeUndefined();
		const fields = runsOf(state).flatMap((run) => (run.formField ? [run.formField] : []));
		expect(fields[1]?.checkBox?.checked).toBe(true);
		expect(fields[2]?.dropDown?.result).toBe(1);
	});

	it('does not operate a field whose fill-in is disabled', () => {
		const state = stateOf();
		const disabled = { ...approved, enabled: false };
		const pos = markerPos(state, 'checkbox');
		const node = state.doc.nodeAt(pos)!;
		const next = state.apply(
			state.tr.setNodeAttribute(
				pos,
				'format',
				JSON.stringify({ ...JSON.parse(node.attrs.format as string), formField: disabled }),
			),
		);
		expect(toggleFormCheckboxAt(next, pos)).toBeUndefined();
	});

	it('treats runs whose field data changed as changed', () => {
		const begin = (formField: FormField): TextRun[] => [
			{ text: '', fieldChar: 'begin', formField },
		];
		expect(sameRuns(begin(approved), begin(approved))).toBe(true);
		expect(
			sameRuns(
				begin(approved),
				begin({ ...approved, checkBox: { default: false, checked: true } }),
			),
		).toBe(false);
	});

	it('finds the result of each text field', () => {
		const [result] = formTextResults(stateOf().doc);
		expect(result).toMatchObject({ field: { name: 'Name' }, length: 3 });
	});

	it('stops typing at the maximum length of a text field', () => {
		const state = stateOf();
		const [result] = formTextResults(state.doc);
		const longer = state.apply(state.tr.insertText('!!!', result!.to));
		expect(longer.doc.textContent).toContain('Ada!!!');
		const tooLong = longer.apply(longer.tr.insertText('!', result!.to + 3));
		expect(tooLong.doc.textContent).toBe(longer.doc.textContent);
	});

	it('keeps a disabled text field read only in a document protected for forms', () => {
		const runs: TextRun[] = [{ text: 'Name: ' }, ...field({ ...name, enabled: false }, 'Ada')];
		const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
		const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
		const forms = { edit: 'forms', enforced: true } as const;
		const state = EditorState.create({ schema, doc, plugins: [formFieldGuardPlugin(() => forms)] });
		const [result] = formTextResults(state.doc);
		expect(state.apply(state.tr.insertText('B', result!.to)).doc.textContent).toBe('Name: Ada');
	});

	it('follows a text field whose result runs on into the next paragraph', () => {
		const first = runToInlineNodes({ text: '', fieldChar: 'begin', formField: name }, schema)
			.concat(runToInlineNodes({ text: '', fieldCode: ' FORMTEXT ' }, schema))
			.concat(runToInlineNodes({ text: '', fieldChar: 'separate' }, schema))
			.concat(runToInlineNodes({ text: 'one', field: { instr: 'FORMTEXT' } }, schema));
		const second = runToInlineNodes({ text: 'two', field: { instr: 'FORMTEXT' } }, schema).concat(
			runToInlineNodes({ text: '', fieldChar: 'end' }, schema),
		);
		const doc = schema.node('doc', null, [
			schema.node('paragraph', null, first),
			schema.node('paragraph', null, second),
		]);
		const [result] = formTextResults(doc);
		expect(result).toMatchObject({ field: { name: 'Name' }, length: 6 });
		expect(doc.textBetween(result!.from, result!.to, '|')).toBe('one|two');
	});

	it('limits the field that grew even when another field disappears in the same edit', () => {
		const runs: TextRun[] = [
			// Longer than the grown field, which pairing fields by their order would compare against.
			...field({ kind: 'text', name: 'First' }, 'abcdefghijkl'),
			{ text: ' and ' },
			...field(name, 'Ada'),
		];
		const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
		const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
		const state = EditorState.create({
			schema,
			doc,
			plugins: [formFieldGuardPlugin(() => undefined)],
		});
		const [first, second] = formTextResults(state.doc);
		const tr = state.tr.insertText('LONGER', second!.to).delete(0, first!.to + 2);
		expect(state.apply(tr).doc.textContent).toBe(state.doc.textContent);
	});

	/** A form protected for filling in whose second section (`p1`) is left unprotected. */
	function sectionedState() {
		const sectioned = new Schema({
			nodes: {
				doc: { content: 'paragraph+', attrs: { sections: { default: null } } },
				paragraph: {
					content: 'inline*',
					attrs: { id: { default: '' }, align: { default: null } },
					toDOM: () => ['p', 0],
				},
				text: { group: 'inline' },
				fieldMarker: fieldMarkerNodeSpec,
			},
			marks: markSpecs,
		});
		const paragraph = (id: string, runs: TextRun[]) =>
			sectioned.node(
				'paragraph',
				{ id },
				runs.flatMap((run) => runToInlineNodes(run, sectioned)),
			);
		const sections = [{ endsAtBlockId: 'p0' }, { endsAtBlockId: 'p1', formProtection: false }];
		const doc = sectioned.node('doc', { sections: JSON.stringify(sections) }, [
			paragraph('p0', [{ text: 'Name: ' }, ...field(name, 'Ada')]),
			paragraph('p1', [{ text: 'Notes' }]),
		]);
		return EditorState.create({
			schema: sectioned,
			doc,
			plugins: [formFieldGuardPlugin(() => ({ edit: 'forms', enforced: true }))],
		});
	}

	it('lets the sections a protected form leaves unprotected change freely', () => {
		const state = sectionedState();
		const notes = state.doc.firstChild!.nodeSize + 1;
		expect(state.apply(state.tr.insertText('My ', notes)).doc.textContent).toContain('My Notes');
		expect(state.apply(state.tr.insertText('X', 1)).doc.textContent).toBe(state.doc.textContent);
		// Joining the free section to the protected one before it is refused.
		const joined = state.apply(state.tr.delete(notes - 2, notes));
		expect(joined.doc.childCount).toBe(2);
	});

	it('lets attributes change in an unprotected section of a protected form, and nowhere else', () => {
		const state = sectionedState();
		const notes = state.doc.firstChild!.nodeSize;
		const aligned = (at: number) => state.apply(state.tr.setNodeAttribute(at, 'align', 'center'));
		expect(aligned(notes).doc.child(1).attrs.align).toBe('center');
		expect(aligned(0).doc.child(0).attrs.align).toBeNull();
	});

	it('only lets form fields change in a document protected for forms', () => {
		const state = stateOf({ edit: 'forms', enforced: true });
		const [result] = formTextResults(state.doc);
		const label = state.apply(state.tr.insertText('X', 1));
		expect(label.doc.textContent).toBe(state.doc.textContent);
		const filled = state.apply(state.tr.insertText('B', result!.to));
		expect(filled.doc.textContent).toContain('AdaB');
		const toggled = state.apply(toggleFormCheckboxAt(state, markerPos(state, 'checkbox'))!);
		expect(
			runsOf(toggled).find((run) => run.formField?.kind === 'checkbox')?.formField?.checkBox,
		).toMatchObject({ checked: true });
		const open = stateOf({ edit: 'forms', enforced: false });
		expect(open.apply(open.tr.insertText('X', 1)).doc.textContent).toMatch(/^XName/);
	});
});

describe('legacy form fields inside locked content controls', () => {
	const control = (
		level: 'inline' | 'block',
		lock: NonNullable<ContentControl['lock']>,
	): ContentControl => ({
		id: '7',
		level,
		kind: 'richText',
		lock,
	});
	function lockedState(around: ContentControl) {
		const runs: TextRun[] = [...field(approved), { text: ' ' }, ...field(status)].map((run) =>
			around.level === 'inline' ? { ...run, contentControl: around } : run,
		);
		const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
		const attrs = around.level === 'block' ? { contentControl: around } : null;
		const doc = schema.node('doc', null, [schema.node('paragraph', attrs, nodes)]);
		return EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
	}

	it('does not operate a check box or drop-down whose control locks its content', () => {
		for (const level of ['inline', 'block'] as const) {
			const state = lockedState(control(level, 'contentLocked'));
			const checkbox = markerPos(state, 'checkbox');
			expect(formFieldLocked(state, checkbox)).toBe(true);
			expect(toggleFormCheckboxAt(state, checkbox)).toBeUndefined();
			expect(selectFormDropDownAt(state, markerPos(state, 'dropDown'), 1)).toBeUndefined();
			// Changing the field data some other way is refused by the lock guard too.
			const format = JSON.parse(state.doc.nodeAt(checkbox)!.attrs.format as string);
			const forced = state.tr.setNodeAttribute(
				checkbox,
				'format',
				JSON.stringify({ ...format, formField: setFormCheckbox(approved, true) }),
			);
			expect(violatesControlLock(forced, state)).toBe(true);
			expect(state.apply(forced).doc.eq(state.doc)).toBe(true);
		}
	});

	it('operates fields in a control that only locks its own removal', () => {
		const state = lockedState(control('block', 'sdtLocked'));
		expect(formFieldLocked(state, markerPos(state, 'checkbox'))).toBe(false);
		const toggled = state.apply(toggleFormCheckboxAt(state, markerPos(state, 'checkbox'))!);
		expect(
			formFieldOfNode(toggled.doc.nodeAt(markerPos(toggled, 'checkbox')))?.checkBox,
		).toMatchObject({
			checked: true,
		});
	});

	it('keeps the field data it does not model through a toggle in the editor', () => {
		const sourceXml = '<w:ffData><w:name w:val="Approved"/><w:tabIndex w:val="4"/></w:ffData>';
		const state = stateOf();
		const pos = markerPos(state, 'checkbox');
		const format = JSON.parse(state.doc.nodeAt(pos)!.attrs.format as string);
		const withSource = state.apply(
			state.tr.setNodeAttribute(
				pos,
				'format',
				JSON.stringify({ ...format, formField: { ...approved, sourceXml } }),
			),
		);
		const toggled = withSource.apply(toggleFormCheckboxAt(withSource, pos)!);
		const begin = runsOf(toggled).find((run) => run.formField?.kind === 'checkbox');
		expect(begin?.formField).toMatchObject({ sourceXml, checkBox: { checked: true } });
	});
});
