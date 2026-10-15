// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
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
function setup(count = 1) {
	vi.useFakeTimers();
	const paper = document.createElement('div'),
		pages = document.createElement('div');
	document.body.append(paper, pages);
	pages.innerHTML = `<div class="dve-print-page" data-page-index="0">${Array.from({ length: count }, (_, index) => `<div class="dve-print-line" data-block-id="${index ? 'body' + index : 'body'}" data-source-start="0" data-source-end="4"><span data-source-start="0" data-source-end="4">Text</span></div>`).join('')}</div>`;
	const line = pages.querySelector<HTMLElement>('.dve-print-line')!,
		span = line.querySelector('span')!;
	vi.spyOn(line, 'getBoundingClientRect').mockReturnValue(new DOMRect(100, 20, 80, 20));
	vi.spyOn(span, 'getBoundingClientRect').mockReturnValue(new DOMRect(100, 20, 40, 20));
	const view = new EditorView(paper, {
		state: EditorState.create({
			doc: schema.node(
				'doc',
				null,
				Array.from({ length: count }, (_, index) =>
					schema.node('paragraph', { id: index ? 'body' + index : 'body' }, schema.text('Text')),
				),
			),
		}),
	});
	vi.spyOn(view, 'coordsAtPos').mockReturnValue({ left: 100, right: 100, top: 20, bottom: 40 });
	vi.spyOn(view, 'hasFocus').mockReturnValue(true);
	const bridge = createPrintEditingBridge(
		pages,
		paper,
		() => view,
		(target) => ({
			blockId: target.closest<HTMLElement>('.dve-print-line')!.dataset.blockId!,
			offset: 1,
		}),
	);
	bridge.setActive(true);
	return { bridge, paper, pages, view };
}
describe('ProseMirror print input bridge', () => {
	it('places the caret on the nearest line when the pointer is off any text', () => {
		const { bridge, pages, view } = setup(1);
		const event = new MouseEvent('pointerdown', {
			bubbles: true,
			cancelable: true,
			button: 0,
			clientX: 400,
			clientY: 300,
		});
		pages.querySelector('.dve-print-page')!.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(true);
		expect(view.state.selection.head).toBe(2);
		bridge.destroy();
		view.destroy();
	});

	it('uses the existing editor selection and rejects unknown stories', () => {
		const { bridge, view } = setup();
		expect(bridge.select('body', 2)).toBe(true);
		expect(view.state.selection.head).toBe(3);
		expect(bridge.select('footnote', 0)).toBe(false);
		expect(view.state.selection.head).toBe(3);
		bridge.destroy();
		view.destroy();
	});
	it('projects selected source ranges and restores the continuous surface', () => {
		const { bridge, paper, pages, view } = setup();
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2, 4)));
		bridge.refresh();
		vi.advanceTimersByTime(20);
		expect(pages.querySelector<HTMLElement>('.dve-print-selection')!.style.width).toBe('20px');
		expect(paper.classList.contains('dve-print-input')).toBe(true);
		bridge.setActive(false);
		vi.advanceTimersByTime(20);
		expect(paper.classList.contains('dve-print-input')).toBe(false);
		expect(pages.querySelector('.dve-print-selection-layer')).toBeNull();
		expect(view.state.doc.textContent).toBe('Text');
		bridge.destroy();
		view.destroy();
	});
	it('does not relocate native input while the browser owns a composition', () => {
		const { bridge, paper, view } = setup();
		Object.defineProperty(view, 'composing', { get: () => true });
		bridge.select('body', 2);
		vi.advanceTimersByTime(20);
		expect(paper.style.translate).toBe('');
		bridge.destroy();
		view.destroy();
	});
	it('indexes a thousand paragraphs once across caret moves and invalidates edited document positions', () => {
		const { bridge, paper, pages, view } = setup(1000);
		const scan = vi.spyOn(view.state.doc, 'descendants');
		const query = vi.spyOn(pages, 'querySelectorAll');
		for (let i = 0; i < 200; i++) expect(bridge.select('body999', i % 4)).toBe(true);
		bridge.refresh();
		vi.advanceTimersByTime(20);
		expect(scan).toHaveBeenCalledTimes(1);
		expect(query).toHaveBeenCalledTimes(1);
		view.dispatch(view.state.tr.insertText('XX', 1));
		const nextScan = vi.spyOn(view.state.doc, 'descendants');
		expect(bridge.select('body999', 2)).toBe(true);
		expect(view.state.selection.head).toBe(5999);
		paper.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true, cancelable: true }),
		);
		expect(view.state.selection.head).toBe(5992);
		expect(nextScan).toHaveBeenCalledTimes(1);
		expect(query).toHaveBeenCalledTimes(1);
		bridge.destroy();
		view.destroy();
	});
	it('disposes native input and global drag listeners before pending drawing can run', () => {
		const { bridge, paper, pages, view } = setup();
		const line = pages.querySelector('.dve-print-line')!;
		Object.defineProperty(document, 'elementFromPoint', { configurable: true, value: () => line });
		const add = vi.spyOn(window, 'addEventListener'),
			remove = vi.spyOn(window, 'removeEventListener');
		const cancel = vi.spyOn(globalThis, 'cancelAnimationFrame');
		pages.dispatchEvent(
			new MouseEvent('pointerdown', { button: 0, bubbles: true, cancelable: true }),
		);
		expect(add).toHaveBeenCalledWith('pointermove', expect.any(Function));
		const dispatch = vi.spyOn(view, 'dispatch');
		bridge.destroy();
		bridge.destroy();
		expect(remove).toHaveBeenCalledWith('pointermove', expect.any(Function));
		expect(remove).toHaveBeenCalledWith('pointerup', expect.any(Function));
		expect(remove).toHaveBeenCalledWith('pointercancel', expect.any(Function));
		expect(cancel).toHaveBeenCalled();
		window.dispatchEvent(new MouseEvent('pointermove', { bubbles: true }));
		pages.dispatchEvent(new MouseEvent('pointerdown', { button: 0, bubbles: true }));
		paper.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
		bridge.refresh();
		bridge.setActive(true);
		expect(bridge.select('body', 2)).toBe(false);
		vi.advanceTimersByTime(20);
		expect(dispatch).not.toHaveBeenCalled();
		expect(paper.classList.contains('dve-print-input')).toBe(false);
		expect(pages.querySelector('.dve-print-selection-layer')).toBeNull();
		delete (document as unknown as { elementFromPoint?: unknown }).elementFromPoint;
		view.destroy();
	});
});
