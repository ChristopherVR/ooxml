import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { ContentControl, TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { runToInlineNodes } from './run-adapter';
import {
	chooseListItem,
	deleteToPlaceholder,
	typeOverPlaceholder,
} from './content-control-commands';
import { contentControlPlugin } from './content-control-guard';
import { controlSpans } from './content-control-spans';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';
import { imageNodeSpec } from './inline-content-schema';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: {
			content: 'inline*',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['p', 0],
		},
		text: { group: 'inline' },
		image: imageNodeSpec,
	},
	marks: markSpecs,
});
const temporary = (extra: Partial<ContentControl> = {}): ContentControl => ({
	id: '21',
	level: 'inline',
	kind: 'richText',
	temporary: true,
	...extra,
});
const stateOf = (runs: TextRun[], block?: ContentControl) => {
	const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
	const paragraph = schema.node('paragraph', block ? { contentControl: block } : null, nodes);
	return EditorState.create({
		schema,
		doc: schema.node('doc', null, [paragraph]),
		plugins: [contentControlPlugin()],
	});
};
const list = temporary({
	kind: 'dropDownList',
	showingPlaceholder: true,
	items: [
		{ display: 'Primary', value: 'primary' },
		{ display: 'Secondary', value: 'secondary' },
	],
});

describe('temporary content controls in the editor', () => {
	it('keeps a picked entry as text and removes the control', () => {
		const state = stateOf([{ text: 'See: ' }, { text: 'Choose an item.', contentControl: list }]);
		const next = state.apply(chooseListItem(state, '21', 1)!);
		expect(next.doc.textContent).toBe('See: Secondary');
		expect(controlSpans(next.doc)).toEqual([]);
	});

	it('keeps text typed over the placeholder and removes the control', () => {
		const control = temporary({ showingPlaceholder: true });
		const state = stateOf([{ text: 'Click here', contentControl: control }]);
		const tr = typeOverPlaceholder(state, 3, 'D');
		const next = state.apply(tr as Exclude<typeof tr, 'blocked' | undefined>);
		expect(next.doc.textContent).toBe('D');
		expect(controlSpans(next.doc)).toEqual([]);
	});

	it('removes the control once its text is edited in any other way', () => {
		const state = stateOf([{ text: 'Draft', contentControl: temporary() }, { text: ' after' }]);
		const next = state.apply(state.tr.insertText('ed', 6));
		expect(next.doc.textContent).toBe('Drafted after');
		expect(controlSpans(next.doc)).toEqual([]);
		const block = stateOf([{ text: 'Notes' }], temporary({ level: 'block' }));
		const edited = block.apply(block.tr.insertText('!', 6));
		expect(edited.doc.firstChild!.attrs.contentControl).toBeNull();
	});

	it('removes the control when a picture in it is replaced, and not when its text is formatted', () => {
		const picture = (partName: string) => schema.node('image', { partName });
		const paragraph = schema.node('paragraph', { contentControl: temporary({ level: 'block' }) }, [
			schema.text('Logo '),
			picture('word/media/image1.png'),
		]);
		const state = EditorState.create({
			schema,
			doc: schema.node('doc', null, [paragraph]),
			plugins: [contentControlPlugin()],
		});
		const replaced = state.apply(state.tr.replaceWith(6, 7, picture('word/media/image2.png')));
		expect(replaced.doc.firstChild!.attrs.contentControl).toBeNull();
		const bold = state.apply(state.tr.addMark(1, 5, schema.marks.bold!.create()));
		expect(bold.doc.firstChild!.attrs.contentControl).toMatchObject({ id: '21' });
	});

	it('removes the text of a temporary control instead of bringing its placeholder back', () => {
		const state = stateOf([{ text: 'x' }, { text: 'Draft', contentControl: temporary() }]);
		const select = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 7)));
		expect(deleteToPlaceholder(select, 'backward')).toBeUndefined();
		const next = select.apply(select.tr.deleteSelection());
		expect(next.doc.textContent).toBe('x');
	});

	it('edits a temporary control that cannot be deleted as an ordinary control', () => {
		const pick = stateOf([
			{ text: 'See: ' },
			{ text: 'Choose an item.', contentControl: { ...list, lock: 'sdtLocked' } },
		]);
		const picked = pick.apply(chooseListItem(pick, '21', 1)!);
		expect(picked.doc.textContent).toBe('See: Secondary');
		expect(controlSpans(picked.doc)).toHaveLength(1);
		const placeholder = temporary({ showingPlaceholder: true, lock: 'sdtLocked' });
		const typing = stateOf([{ text: 'Click here', contentControl: placeholder }]);
		const typed = typeOverPlaceholder(typing, 3, 'D');
		const value = typing.apply(typed as Exclude<typeof typed, 'blocked' | undefined>);
		expect(value.doc.textContent).toBe('D');
		expect(controlSpans(value.doc)).toHaveLength(1);
		const locked = temporary({ kind: 'text', lock: 'sdtLocked' });
		const text = stateOf([{ text: 'x' }, { text: 'Draft', contentControl: locked }]);
		const all = text.apply(text.tr.setSelection(TextSelection.create(text.doc, 2, 7)));
		const restored = all.apply(deleteToPlaceholder(all, 'backward')!);
		expect(restored.doc.textContent).toBe('xClick or tap here to enter text.');
		expect(controlSpans(restored.doc)[0]?.control.showingPlaceholder).toBe(true);
	});

	it("leaves controls that share a temporary control's id alone when it is not edited", () => {
		// A pasted copy keeps an id in other tools, and some files repeat ids: the copy here is no
		// longer temporary, and the other two temporary controls hold different text.
		const state = stateOf([
			{ text: 'Draft', contentControl: temporary() },
			{ text: ' and ' },
			{ text: 'Kept', contentControl: { ...temporary(), temporary: false } },
			{ text: ' and ' },
			{ text: 'Second', contentControl: temporary() },
		]);
		const next = state.apply(state.tr.insertText('!', 7));
		expect(next.doc.textContent).toBe('Draft !and Kept and Second');
		expect(controlSpans(next.doc).map((span) => span.control.id)).toEqual(['21', '21', '21']);
	});

	it('takes a block control off all of its paragraphs, also around an inline control', () => {
		const block = temporary({ id: '31', level: 'block' });
		const box: ContentControl = { id: '32', level: 'inline', kind: 'richText' };
		const doc = schema.node('doc', null, [
			schema.node('paragraph', { contentControl: block }, [
				schema.text('Alpha '),
				schema.text('box', [schema.marks.contentControl!.create({ data: JSON.stringify(box) })]),
			]),
			schema.node('paragraph', { contentControl: block }, [schema.text('Beta')]),
			schema.node('paragraph', null, [schema.text('Elsewhere')]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const blocks = (s: EditorState) => {
			const controls: unknown[] = [];
			s.doc.forEach((node) => controls.push(node.attrs.contentControl?.id ?? null));
			return controls;
		};
		// An edit outside the control leaves it whole.
		const outside = state.apply(state.tr.insertText('!', state.doc.content.size - 1));
		expect(blocks(outside)).toEqual(['31', '31', null]);
		// Editing its second paragraph removes it from the first too.
		const edited = state.apply(state.tr.insertText('!', 14));
		expect(edited.doc.child(1).textContent).toBe('Be!ta');
		expect(blocks(edited)).toEqual([null, null, null]);
	});

	it('leaves a control that cannot be deleted, and edits from elsewhere, as they are', () => {
		const locked = stateOf([{ text: 'Draft', contentControl: temporary({ lock: 'sdtLocked' }) }]);
		expect(controlSpans(locked.apply(locked.tr.insertText('s', 6)).doc)).toHaveLength(1);
		const shared = stateOf([{ text: 'Draft', contentControl: temporary() }]);
		const remote = shared.tr.insertText('s', 6).setMeta(REMOTE_TRANSACTION_META, true);
		expect(controlSpans(shared.apply(remote).doc)).toHaveLength(1);
		const beside = stateOf([{ text: 'Before ' }, { text: 'Draft', contentControl: temporary() }]);
		const untouched = beside.apply(beside.tr.insertText('e', 2));
		expect(untouched.doc.textContent).toBe('Beefore Draft');
		expect(controlSpans(untouched.doc)).toHaveLength(1);
	});
});
