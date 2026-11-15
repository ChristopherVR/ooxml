// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { createDocument } from 'ooxml-core/docx';
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { modelToDoc } from './model-adapter';
import {
	applyPasteOption,
	dismissPasteOptions,
	PasteController,
	pastedRange,
	pasteOptionsPlugin,
	pastePayload,
} from './paste-options';
import { pasteOptionItems, pasteOptionsButtonPlugin } from './paste-options-button';
import { wordPastePlugin } from './word-paste';

const fixture = (name: string): string =>
	readFileSync(
		path.join(import.meta.dirname, `../../../core/docx/__fixtures__/clipboard/${name}.html`),
		'utf8',
	);

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.innerHTML = '';
	vi.unstubAllGlobals();
});

function editor(withButton = false) {
	vi.stubGlobal('ClipboardEvent', Event);
	const model = createDocument();
	const controller = new PasteController();
	const container = document.createElement('div');
	const host = document.createElement('div');
	document.body.append(container, host);
	const view = new EditorView(host, {
		handleScrollToSelection: () => true,
		state: EditorState.create({
			doc: modelToDoc(model),
			plugins: [
				history(),
				wordPastePlugin(
					() => model,
					() => (controller.option === 'merge' ? 'merge' : 'source'),
				),
				pasteOptionsPlugin(controller),
				...(withButton
					? [pasteOptionsButtonPlugin({ container: () => container, locale: () => 'en' })]
					: []),
			],
		}),
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
	views.push(view);
	return { view, container };
}

const payload = (name: string) => ({ html: fixture(name), text: 'Plain words' });
const paragraphs = (view: EditorView) => {
	const found: Array<{ text: string; numId: unknown; style: unknown; marks: string[] }> = [];
	view.state.doc.descendants((node) => {
		if (node.type.name !== 'paragraph') return;
		const marks: string[] = [];
		node.descendants((child) => {
			for (const mark of child.marks) marks.push(mark.type.name);
		});
		found.push({ text: node.textContent, numId: node.attrs.numId, style: node.attrs.style, marks });
	});
	return found;
};

describe('Paste Options', () => {
	it('remembers the pasted range and what the clipboard held', () => {
		const { view } = editor();
		expect(pastePayload(view, payload('word-numbered'), 'source')).toBe(true);
		const range = pastedRange(view)!;
		expect(range.option).toBe('source');
		expect(view.state.doc.textBetween(range.from, range.to, '\n')).toContain('Step one');
		expect(view.state.doc.textBetween(range.from, range.to, '\n')).toContain('Step two');
	});

	it('keeps text only, as one change that undoes back to the formatted paste', () => {
		const { view } = editor();
		pastePayload(view, payload('word-numbered'), 'source');
		const formatted = view.state.doc;
		expect(applyPasteOption(view, 'text')).toBe(true);
		expect(paragraphs(view).map((p) => p.text)).toEqual(['Plain words']);
		expect(pastedRange(view)?.option).toBe('text');
		undo(view.state, view.dispatch);
		expect(view.state.doc.eq(formatted)).toBe(true);
	});

	it('merges formatting: lists and emphasis stay, styles and run properties go', () => {
		const { view } = editor();
		pastePayload(view, payload('word-styles'), 'source');
		expect(paragraphs(view).find((p) => p.text === 'Clipboard Heading')?.style).toBe('Heading1');
		applyPasteOption(view, 'merge');
		const merged = paragraphs(view);
		expect(merged.every((p) => !p.style)).toBe(true);
		const plain = merged.find((p) => p.text.startsWith('Plain then'))!;
		expect(plain.marks).toContain('bold');
		expect(plain.marks).not.toContain('runProperties');
		expect(plain.marks).not.toContain('font');
		applyPasteOption(view, 'source');
		pastePayload(view, payload('word-numbered'), 'source');
		applyPasteOption(view, 'merge');
		expect(paragraphs(view).filter((p) => p.text.startsWith('Step'))[0]?.numId).not.toBeNull();
	});

	it('ends the offer on typing or Escape, not on attribute repairs', () => {
		const { view } = editor();
		pastePayload(view, payload('word-styles'), 'source');
		const range = pastedRange(view)!;
		view.dispatch(view.state.tr.setNodeAttribute(range.from, 'id', 'repaired'));
		expect(pastedRange(view)).not.toBeNull();
		view.dispatch(view.state.tr.insertText('x'));
		expect(pastedRange(view)).toBeNull();
		pastePayload(view, payload('word-styles'), 'source');
		expect(dismissPasteOptions(view)).toBe(true);
		expect(pastedRange(view)).toBeNull();
	});

	it('offers only text when the clipboard held no HTML', () => {
		const { view } = editor();
		pastePayload(view, { html: null, text: 'just text' }, 'text');
		const items = pasteOptionItems(pastedRange(view)!, 'en');
		expect(items.map((item) => [item.label, item.disabled, item.checked])).toEqual([
			['Keep Source Formatting', true, false],
			['Merge Formatting', true, false],
			['Keep Text Only', false, true],
		]);
		expect(applyPasteOption(view, 'merge')).toBe(false);
	});

	it('shows a button after a paste whose menu pastes again', () => {
		const { view, container } = editor(true);
		pastePayload(view, payload('word-numbered'), 'source');
		const button = container.querySelector<HTMLButtonElement>('.dve-paste-options')!;
		expect(button.getAttribute('aria-label')).toBe('Paste Options (Ctrl)');
		button.click();
		const menu = container.querySelector('office-ui-context-menu')!;
		menu.dispatchEvent(new CustomEvent('office-menu-request', { detail: { id: 'text' } }));
		expect(paragraphs(view).map((p) => p.text)).toEqual(['Plain words']);
		view.dispatch(view.state.tr.insertText('typed'));
		expect(container.querySelector('.dve-paste-options')).toBeNull();
	});
});
