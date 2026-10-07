// @vitest-environment jsdom
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { fieldGuardPlugin, fieldResultRanges, runToInlineNodes } from 'ooxml-core/docx/ui';
import type { TextRun } from 'ooxml-core/docx';
import { schema } from './schema';

const views: EditorView[] = [];
beforeAll(() => {
	Object.assign(Range.prototype, {
		getClientRects: () => [],
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.innerHTML = '';
	vi.unstubAllGlobals();
});
function editor(simple: boolean) {
	vi.stubGlobal('ClipboardEvent', Event);
	const result: TextRun = {
		text: 'ABCDE',
		bold: true,
		field: { instr: 'REF Bookmark', ...(simple ? { simple: true } : {}) },
		...(simple ? { fieldInstanceId: 'first' } : {}),
	};
	const runs: TextRun[] = simple
		? [result]
		: [
				{ text: '', fieldChar: 'begin' },
				{ text: '', fieldCode: ' REF Bookmark ' },
				{ text: '', fieldChar: 'separate' },
				result,
				{ text: '', fieldChar: 'end' },
			];
	const doc = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			schema.text('L'),
			...runs.flatMap((run) => runToInlineNodes(run, schema)),
			schema.text('R'),
		]),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({ doc, plugins: [history(), fieldGuardPlugin()] }),
	});
	views.push(view);
	return view;
}
const replacements = [
	{ name: 'partial', from: 1, to: 2, expected: 'AXCDE' },
	{ name: 'whole', from: 0, to: 5, expected: 'X' },
	{ name: 'start', from: 0, to: 1, expected: 'XBCDE' },
	{ name: 'end', from: 4, to: 5, expected: 'ABCDX' },
];
for (const simple of [true, false])
	for (const replacement of replacements)
		it(`retains a ${simple ? 'simple' : 'complex'} field during ${replacement.name} result paste`, () => {
			const view = editor(simple);
			const original = view.state.doc;
			const field = fieldResultRanges(original)[0]!;
			view.dispatch(
				view.state.tr.setSelection(
					TextSelection.create(
						view.state.doc,
						field.from + replacement.from,
						field.from + replacement.to,
					),
				),
			);
			expect(view.pasteHTML('<b>X</b>')).toBe(true);
			expect(fieldResultRanges(view.state.doc).map((run) => run.text)).toEqual([
				replacement.expected,
			]);
			if (simple)
				expect(
					fieldResultRanges(view.state.doc)[0]!.marks.find(
						(mark) => mark.type.name === 'runProperties',
					)!.attrs.props.fieldInstanceId,
				).toBe('first');
			undo(view.state, view.dispatch);
			expect(view.state.doc.eq(original)).toBe(true);
		});
for (const replacement of replacements)
	it(`retains a simple field during ${replacement.name} result typing`, () => {
		const view = editor(true);
		const original = view.state.doc;
		const field = fieldResultRanges(original)[0]!;
		const from = field.from + replacement.from;
		const to = field.from + replacement.to;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
		let handled = false;
		view.someProp('handleTextInput', (handler) => {
			handled = Boolean(
				handler(view, from, to, 'X', () => view.state.tr.insertText('X', from, to)),
			);
			return handled;
		});
		expect(handled).toBe(true);
		expect(fieldResultRanges(view.state.doc).map((run) => run.text)).toEqual([
			replacement.expected,
		]);
		expect(
			fieldResultRanges(view.state.doc)[0]!.marks.find(
				(mark) => mark.type.name === 'runProperties',
			)!.attrs.props.fieldInstanceId,
		).toBe('first');
		undo(view.state, view.dispatch);
		expect(view.state.doc.eq(original)).toBe(true);
	});
for (const simple of [true, false])
	it(`copies ${simple ? 'simple' : 'complex'} field result text as formatted literal text`, () => {
		const view = editor(simple);
		const range = fieldResultRanges(view.state.doc)[0]!;
		view.dispatch(
			view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to)),
		);
		const original = view.state.doc;
		const { dom } = view.serializeForClipboard(view.state.selection.content());
		expect(view.state.doc.eq(original)).toBe(true);
		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, simple ? range.to : range.to + 1),
			),
		);
		expect(view.pasteHTML(dom.innerHTML)).toBe(true);
		expect(view.state.doc.textContent).toBe('LABCDEABCDER');
		expect(fieldResultRanges(view.state.doc).map((run) => run.text)).toEqual(['ABCDE']);
		const copy = view.state.doc.nodeAt(simple ? range.to : range.to + 1)!;
		expect(copy.marks.some((mark) => mark.type.name === 'bold')).toBe(true);
		expect(
			copy.marks.find((mark) => mark.type.name === 'runProperties')?.attrs.props?.fieldInstanceId,
		).toBeUndefined();
		undo(view.state, view.dispatch);
		expect(view.state.doc.eq(original)).toBe(true);
	});

it('copies a complete complex field as a second field and preserves the source through undo', () => {
	const view = editor(false);
	const original = view.state.doc;
	const result = fieldResultRanges(original)[0]!;
	view.updateState(
		EditorState.create({
			doc: original,
			selection: TextSelection.create(original, 2, result.to + 1),
			plugins: [history(), fieldGuardPlugin()],
		}),
	);
	const { dom } = view.serializeForClipboard(view.state.selection.content());
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, result.to + 1)));
	expect(view.pasteHTML(dom.innerHTML)).toBe(true);
	expect(fieldResultRanges(view.state.doc).map((run) => run.text)).toEqual(['ABCDE', 'ABCDE']);
	undo(view.state, view.dispatch);
	expect(view.state.doc.eq(original)).toBe(true);
});
