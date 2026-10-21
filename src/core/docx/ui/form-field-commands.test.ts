import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import type { ContentControl, DocumentProtection, FormField, TextRun } from '../model';
import { setFormCheckbox } from '../form-fields';
import { contentControlPlugin, violatesControlLock } from './content-control-guard';
import { markSpecs } from './schema-marks';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { appendInlineNode, runToInlineNodes } from './run-adapter';
import {
	formFieldBeside,
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
/** Word's result for a text field with no text: five en spaces, which keep the field visible. */
const BLANK = '\u2002'.repeat(5);
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

	it('finds the check box or drop-down beside a position, for keys to operate it', () => {
		const state = stateOf();
		const box = markerPos(state, 'checkbox');
		const list = markerPos(state, 'dropDown');
		// In front of the box, among its hidden markers, and just after it.
		for (const at of [box, box + 1, box + 2, box + 3])
			expect(formFieldBeside(state.doc, at)).toBe(box);
		expect(formFieldBeside(state.doc, box - 1)).toBeUndefined();
		expect(formFieldBeside(state.doc, list)).toBe(list);
		// Not a text field, whose result takes typing.
		const [result] = formTextResults(state.doc);
		expect(formFieldBeside(state.doc, result!.from)).toBeUndefined();
		// Between two fields, the one that starts there.
		const runs: TextRun[] = [...field(approved), ...field(status)];
		const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
		const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
		expect(formFieldBeside(doc, 4)).toBe(4);
		expect(formFieldOfNode(doc.nodeAt(4))?.kind).toBe('dropDown');
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

	it("replaces a blank text field's en spaces with the first text typed into it", () => {
		const code: FormField = { kind: 'text', name: 'Code', textInput: { maxLength: 4 } };
		const runs: TextRun[] = [{ text: 'Code: ' }, ...field(code, BLANK), { text: ' end' }];
		const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
		const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
		let state = EditorState.create({
			schema,
			doc,
			plugins: [formFieldGuardPlugin(() => undefined)],
		});
		const result = () => formTextResults(state.doc)[0]!;
		// An edit elsewhere leaves the blank as it is.
		state = state.apply(state.tr.insertText('Post ', 1));
		expect(state.doc.textContent).toBe(`Post Code: ${BLANK} end`);
		// Typed anywhere in the blank, the first character takes the en spaces' place, and the
		// maximum length counts only what was typed.
		state = state.apply(state.tr.insertText('A', result().from + 2));
		expect(state.doc.textContent).toBe('Post Code: A end');
		for (const character of 'B12') state = state.apply(state.tr.insertText(character, result().to));
		expect(state.doc.textContent).toBe('Post Code: AB12 end');
		state = state.apply(state.tr.insertText('!', result().to));
		expect(state.doc.textContent).toBe('Post Code: AB12 end');
		expect(result().length).toBe(4);
	});

	it('takes text pasted into a blank text field up to its maximum length', () => {
		const code: FormField = { kind: 'text', name: 'Code', textInput: { maxLength: 3 } };
		const runs: TextRun[] = [...field(code, BLANK)];
		const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
		const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
		const state = EditorState.create({
			schema,
			doc,
			plugins: [formFieldGuardPlugin(() => ({ edit: 'forms', enforced: true }))],
		});
		const [result] = formTextResults(state.doc);
		const pasted = (text: string) => state.apply(state.tr.insertText(text, result!.from)).doc;
		expect(pasted('XYZ').textContent).toBe('XYZ');
		expect(pasted('WXYZ').textContent).toBe(BLANK);
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

	it('keeps a disabled text field read only inside a content control of a protected form', () => {
		for (const level of ['inline', 'block'] as const) {
			const control: ContentControl = { id: '8', level, kind: 'richText' };
			const runs: TextRun[] = [
				{ text: 'Ref: ' },
				...field({ ...name, enabled: false }, 'Fixed'),
				{ text: ' note' },
			].map((run) => (level === 'inline' ? { ...run, contentControl: control } : run));
			const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
			const attrs = level === 'block' ? { contentControl: control } : null;
			const doc = schema.node('doc', null, [schema.node('paragraph', attrs, nodes)]);
			const forms = { edit: 'forms', enforced: true } as const;
			const state = EditorState.create({
				schema,
				doc,
				plugins: [formFieldGuardPlugin(() => forms)],
			});
			const [result] = formTextResults(state.doc);
			const text = (tr: Transaction) => state.apply(tr).doc.textContent;
			// The control's own text can change; the disabled field's result cannot, nor can it go.
			expect(text(state.tr.insertText('!', state.doc.content.size - 1))).toBe('Ref: Fixed note!');
			expect(text(state.tr.insertText('B', result!.to))).toBe('Ref: Fixed note');
			expect(text(state.tr.insertText('B', result!.from + 1))).toBe('Ref: Fixed note');
			expect(text(state.tr.delete(result!.from, result!.to))).toBe('Ref: Fixed note');
			expect(text(state.tr.delete(result!.begin, result!.to + 1))).toBe('Ref: Fixed note');
			// Nor its formatting: mark steps map no positions, so they are checked by their range.
			const bold = schema.marks.bold!;
			const bolded = (tr: Transaction, from: number, to: number) =>
				state.apply(tr).doc.rangeHasMark(from, to, bold);
			const end = state.doc.content.size - 1;
			expect(bolded(state.tr.addMark(end - 4, end, bold.create()), end - 4, end)).toBe(true);
			expect(
				bolded(state.tr.addMark(result!.from, result!.to, bold.create()), result!.from, result!.to),
			).toBe(false);
		}
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

	type SectionLayout = { endsAtBlockId: string; formProtection?: false }[];

	/**
	 * A form protected for filling in, by default with a protected first section (`p0`) and a last
	 * section (`p1`) the protection leaves unprotected.
	 */
	function sectionedState(
		sections: SectionLayout = [
			{ endsAtBlockId: 'p0' },
			{ endsAtBlockId: 'p1', formProtection: false },
		],
		texts: Record<string, string> = {},
		first: TextRun[] = [{ text: 'Name: ' }, ...field(name, 'Ada')],
	) {
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
		const extra = Object.entries(texts).map(([id, text]) => paragraph(id, [{ text }]));
		const doc = sectioned.node('doc', { sections: JSON.stringify(sections) }, [
			paragraph('p0', first),
			paragraph('p1', [{ text: 'Notes' }]),
			...extra,
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

	/** The text of each paragraph, after typing `text` at the start of paragraph `index`. */
	const typedAt = (state: EditorState, index: number, text: string): string[] => {
		let pos = 0;
		for (let at = 0; at < index; at++) pos += state.doc.child(at).nodeSize;
		const typed = state.apply(state.tr.insertText(text, pos + 1));
		const texts: string[] = [];
		typed.doc.forEach((node) => texts.push(node.textContent));
		return texts;
	};

	it('lets a paragraph added at the end of an unprotected last section be typed into', () => {
		const state = sectionedState();
		// Enter at the end of the section's last paragraph: the new paragraph has no id yet.
		const end = state.doc.content.size - 1;
		const split = state.apply(state.tr.split(end, 1, [{ type: state.schema.nodes.paragraph! }]));
		expect(split.doc.childCount).toBe(3);
		expect(typedAt(split, 2, 'More')).toEqual(['Name: Ada', 'Notes', 'More']);
	});

	it('keeps the last section editable when the block it named is gone', () => {
		const state = sectionedState([
			{ endsAtBlockId: 'p0' },
			{ endsAtBlockId: 'removed', formProtection: false },
		]);
		expect(typedAt(state, 1, 'My ')).toEqual(['Name: Ada', 'My Notes']);
		expect(typedAt(state, 0, 'X')).toEqual(['Name: Ada', 'Notes']);
	});

	it('gives the blocks of a section whose last block is gone to the next section', () => {
		const state = sectionedState(
			[
				{ endsAtBlockId: 'p0' },
				{ endsAtBlockId: 'removed' },
				{ endsAtBlockId: 'p2', formProtection: false },
			],
			{ p2: 'Remarks' },
		);
		expect(typedAt(state, 1, 'My ')).toEqual(['Name: Ada', 'My Notes', 'Remarks']);
		expect(typedAt(state, 2, 'Our ')).toEqual(['Name: Ada', 'Notes', 'Our Remarks']);
	});

	it('refuses an edit that would move text of an unprotected section into a protected one', () => {
		// The first section is unprotected and ends at `p0`; a split puts its second part in the
		// next section, as the editor and the saved document place a section break on the first.
		const state = sectionedState([
			{ endsAtBlockId: 'p0', formProtection: false },
			{ endsAtBlockId: 'p1' },
		]);
		const end = state.doc.child(0).nodeSize - 1;
		const atEnd = state.apply(state.tr.split(end, 1, [{ type: state.schema.nodes.paragraph! }]));
		expect(atEnd.doc.childCount).toBe(2);
		const inside = state.apply(state.tr.split(3));
		expect(inside.doc.childCount).toBe(2);
		expect(typedAt(state, 0, 'My ')).toEqual(['My Name: Ada', 'Notes']);
	});

	it('lets paragraphs split anywhere else in an unprotected section', () => {
		const state = sectionedState(
			[{ endsAtBlockId: 'p1', formProtection: false }, { endsAtBlockId: 'p2' }],
			{ p2: 'Protected' },
		);
		// `p0` is in the unprotected section but does not end it.
		expect(state.apply(state.tr.split(3)).doc.childCount).toBe(4);
		const start = state.doc.child(0).nodeSize - 1;
		expect(
			state.apply(state.tr.split(start, 1, [{ type: state.schema.nodes.paragraph! }])).doc
				.childCount,
		).toBe(4);
	});

	it('lets a disabled text field change in a section the protection leaves unprotected', () => {
		const state = sectionedState(
			[{ endsAtBlockId: 'p0', formProtection: false }, { endsAtBlockId: 'p1' }],
			{},
			[{ text: 'Name: ' }, ...field({ ...name, enabled: false }, 'Ada')],
		);
		const [result] = formTextResults(state.doc);
		expect(state.apply(state.tr.insertText('B', result!.to)).doc.textContent).toContain('AdaB');
	});

	it('lets the editor give a paragraph another id unless that moves a section', () => {
		// The first section, protected, holds `p0` and ends at `p1`; the last is left unprotected.
		const state = sectionedState(
			[{ endsAtBlockId: 'p1' }, { endsAtBlockId: 'p2', formProtection: false }],
			{ p2: 'Free' },
		);
		const rename = (index: number, attrs: Record<string, unknown>) => {
			let pos = 0;
			for (let at = 0; at < index; at++) pos += state.doc.child(at).nodeSize;
			const node = state.doc.child(index);
			return state
				.apply(state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs }))
				.doc.child(index).attrs;
		};
		expect(rename(0, { id: 'p9' })).toMatchObject({ id: 'p9' });
		expect(rename(0, { id: 'p9', align: 'center' })).toMatchObject({ id: 'p0', align: null });
		// Renaming the block a protected section ends at would give its text to the next section.
		expect(rename(1, { id: 'p9' })).toMatchObject({ id: 'p1' });
		expect(
			state.apply(state.tr.setNodeAttribute(state.doc.child(0).nodeSize, 'id', 'p9')).doc.child(1)
				.attrs.id,
		).toBe('p1');
	});

	it('lets a split move text between two unprotected sections', () => {
		const state = sectionedState([
			{ endsAtBlockId: 'p0', formProtection: false },
			{ endsAtBlockId: 'p1', formProtection: false },
		]);
		const end = state.doc.child(0).nodeSize - 1;
		expect(
			state.apply(state.tr.split(end, 1, [{ type: state.schema.nodes.paragraph! }])).doc.childCount,
		).toBe(3);
		expect(state.apply(state.tr.split(3)).doc.childCount).toBe(3);
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
