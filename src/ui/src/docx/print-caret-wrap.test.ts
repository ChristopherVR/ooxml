// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createPrintEditingBridge } from './print-editing';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'text*', attrs: { id: { default: null } }, toDOM: () => ['p', 0] },
		text: {},
	},
});
afterEach(() => {
	vi.restoreAllMocks();
	vi.useRealTimers();
	document.body.replaceChildren();
});

/** One paragraph "aaaa bbbb" wrapped into two lines that meet at offset 5. */
function wrapped() {
	vi.useFakeTimers();
	const paper = document.createElement('div'),
		pages = document.createElement('div');
	document.body.append(paper, pages);
	pages.innerHTML =
		'<div class="dve-print-page" data-page-index="0">' +
		'<div class="dve-print-line" data-block-id="p" data-source-start="0" data-source-end="5"><span data-source-start="0" data-source-end="5">aaaa </span></div>' +
		'<div class="dve-print-line" data-block-id="p" data-source-start="5" data-source-end="9"><span data-source-start="5" data-source-end="9">bbbb</span></div></div>';
	const lines = [...pages.querySelectorAll<HTMLElement>('.dve-print-line')];
	lines.forEach((line, i) => {
		vi.spyOn(line, 'getBoundingClientRect').mockReturnValue(new DOMRect(100, 20 + i * 20, 80, 20));
		vi.spyOn(line.firstElementChild!, 'getBoundingClientRect').mockReturnValue(
			new DOMRect(100, 20 + i * 20, 40, 20),
		);
	});
	const view = new EditorView(paper, {
		state: EditorState.create({
			doc: schema.node('doc', null, [
				schema.node('paragraph', { id: 'p' }, schema.text('aaaa bbbb')),
			]),
		}),
	});
	vi.spyOn(view, 'coordsAtPos').mockReturnValue({ left: 100, right: 100, top: 20, bottom: 40 });
	vi.spyOn(view, 'hasFocus').mockReturnValue(true);
	const bridge = createPrintEditingBridge(
		pages,
		paper,
		() => view,
		() => null,
	);
	bridge.setActive(true);
	const caretTop = () => {
		vi.runAllTimers();
		return (pages.querySelector('.dve-print-caret') as HTMLElement).style.top;
	};
	return { bridge, view, caretTop };
}

it('draws a boundary caret on the earlier line when upstream and the later one when downstream', () => {
	const { bridge, view, caretTop } = wrapped();
	bridge.select('p', 5, false, 'upstream');
	expect(caretTop()).toBe('20px');
	bridge.select('p', 5, false, 'downstream');
	expect(caretTop()).toBe('40px');
	bridge.select('p', 5);
	expect(caretTop()).toBe('40px');
	bridge.destroy();
	view.destroy();
});
