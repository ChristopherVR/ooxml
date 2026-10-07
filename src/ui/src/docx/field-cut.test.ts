// @vitest-environment jsdom
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import type { TextRun } from 'ooxml-core/docx';
import { fieldGuardPlugin, runToInlineNodes } from 'ooxml-core/docx/ui';
import { schema } from './schema';
import { copyOrCut } from './context-menu-actions';

const views: EditorView[] = [];
beforeAll(() =>
	Object.assign(Range.prototype, {
		getClientRects: () => [],
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	}),
);
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});
function editor(simple: boolean) {
	vi.stubGlobal('ClipboardEvent', Event);
	const result: TextRun = {
		text: 'AB',
		bold: true,
		field: { instr: 'REF Target', ...(simple ? { simple: true } : {}) },
		...(simple ? { fieldInstanceId: 'first' } : {}),
	};
	const runs: TextRun[] = simple
		? [result]
		: [
				{ text: '', fieldChar: 'begin' },
				{ text: '', fieldCode: 'REF Target' },
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
	const from = simple ? 2 : 5;
	const view = new EditorView(document.body.appendChild(document.createElement('div')), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, from, from + 2),
			plugins: [history(), fieldGuardPlugin()],
		}),
		handleScrollToSelection: () => true,
	});
	views.push(view);
	return view;
}
for (const simple of [true, false])
	for (const method of ['event', 'fallback'])
		it(`${method} cut keeps an empty ${simple ? 'simple' : 'complex'} field and undo restores its result`, async () => {
			const view = editor(simple);
			const original = view.state.doc;
			if (method === 'fallback') {
				Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
				const writeText = vi.fn(async (_text: string) => {});
				Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
				expect(await copyOrCut(view, 'cut')).toBe(true);
				expect(writeText).toHaveBeenCalledWith('AB');
			} else {
				const data: Record<string, string> = {};
				const event = new Event('cut', { bubbles: true, cancelable: true });
				Object.defineProperty(event, 'clipboardData', {
					value: {
						clearData: () => {},
						setData: (kind: string, text: string) => (data[kind] = text),
					},
				});
				view.dom.dispatchEvent(event);
				expect(event.defaultPrevented).toBe(true);
				expect(data['text/plain']).toBe('AB');
				expect(data['text/html']).toContain('<strong');
				expect(data['text/html']).not.toContain('data-field');
			}
			expect(view.state.doc.textContent).toBe('LR');
			const codes: string[] = [];
			view.state.doc.descendants((node) => {
				if (node.attrs.kind === 'code') codes.push(node.attrs.code);
			});
			expect(codes).toEqual(['REF Target']);
			undo(view.state, view.dispatch);
			expect(view.state.doc.eq(original)).toBe(true);
		});

it('leaves a selected field intact when fallback clipboard write is denied', async () => {
	const view = editor(true);
	const original = view.state.doc;
	Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
	Object.defineProperty(navigator, 'clipboard', {
		configurable: true,
		value: {
			writeText: async () => {
				throw new Error('denied');
			},
		},
	});
	expect(await copyOrCut(view, 'cut')).toBe(false);
	expect(view.state.doc.eq(original)).toBe(true);
});
it('does not handle a field cut event in read-only mode', () => {
	const view = editor(true);
	const original = view.state.doc;
	view.setProps({ editable: () => false });
	const setData = vi.fn();
	const event = new Event('cut', { bubbles: true, cancelable: true });
	Object.defineProperty(event, 'clipboardData', { value: { clearData: () => {}, setData } });
	view.dom.dispatchEvent(event);
	expect(setData).not.toHaveBeenCalled();
	expect(view.state.doc.eq(original)).toBe(true);
});
