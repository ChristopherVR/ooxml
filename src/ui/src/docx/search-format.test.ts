// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument } from 'ooxml-core/docx';
import { afterEach, describe, expect, it } from 'vitest';
import { docToModel, modelToDoc } from './model-adapter';
import { createSearchPanel } from './search-panel';

afterEach(() => {
	document.body.innerHTML = '';
});

function mountEditor(blocks: ReturnType<typeof createDocument>['blocks']) {
	const model = createDocument();
	model.blocks = blocks;
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({ doc: modelToDoc(model) }),
	});
	const panel = createSearchPanel({ getView: () => view });
	document.body.append(panel.element);
	panel.open();
	return { view, panel, model };
}

function field<T extends HTMLElement>(
	root: HTMLElement,
	kind: 'find' | 'replace',
	name: string,
): T {
	return root.querySelector<T>(`[data-format-kind="${kind}"] [data-format-field="${name}"]`)!;
}

describe('Word Find & Replace Format > Font, Paragraph, and Style UI pickers', () => {
	it('filters text matches by Find format (Font) and replaces with Replace format', () => {
		const { view, panel, model } = mountEditor([
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					{ text: 'hello ', bold: false },
					{ text: 'hello', bold: true },
				],
			},
		]);

		const findInput = panel.element.querySelector<HTMLInputElement>('input[type="search"]')!;
		const replaceInput = panel.element.querySelector<HTMLInputElement>(
			'input[aria-label="Replace with"]',
		)!;
		const status = panel.element.querySelector<HTMLElement>('[role="status"]')!;
		const findBold = field<HTMLSelectElement>(panel.element, 'find', 'bold');
		const replaceItalic = field<HTMLSelectElement>(panel.element, 'replace', 'italic');
		const replaceColor = field<HTMLInputElement>(panel.element, 'replace', 'color');
		const replaceAllBtn = panel.element.querySelector<HTMLButtonElement>(
			'button[aria-label="Replace all"]',
		)!;

		findInput.value = 'hello';
		findInput.dispatchEvent(new Event('input'));
		expect(status.textContent).toContain('2');

		// Filter to Bold only -> 1 match
		findBold.value = 'on';
		findBold.dispatchEvent(new Event('change'));
		expect(status.textContent).toContain('1');
		expect(panel.element.querySelector('[data-format-summary="find"]')?.textContent).toContain(
			'Bold',
		);

		// Replace matched bold "hello" with italic #CC0000 "world"
		replaceInput.value = 'world';
		replaceItalic.value = 'on';
		replaceItalic.dispatchEvent(new Event('change'));
		replaceColor.value = '#CC0000';
		replaceColor.dispatchEvent(new Event('input'));
		replaceAllBtn.click();

		const saved = docToModel(view.state.doc, model);
		const p1 = saved.blocks[0]!;
		expect(p1.type).toBe('paragraph');
		if (p1.type === 'paragraph') {
			expect(p1.runs.map((r) => r.text).join('')).toBe('hello world');
			const world = p1.runs.find((r) => r.text === 'world');
			expect(world?.bold).toBe(true);
			expect(world?.italic).toBe(true);
			expect(world?.color).toBe('#CC0000');
		}
		view.destroy();
	});

	it('supports format-only search and format-only replacement for Paragraph alignment and Style', () => {
		const { view, panel, model } = mountEditor([
			{
				type: 'paragraph',
				id: 'p1',
				style: 'Heading1',
				align: 'left',
				runs: [{ text: 'Chapter One' }],
			},
			{
				type: 'paragraph',
				id: 'p2',
				style: 'Normal',
				align: 'left',
				runs: [{ text: 'Body text' }],
			},
		]);

		const status = panel.element.querySelector<HTMLElement>('[role="status"]')!;
		const findStyle = field<HTMLSelectElement>(panel.element, 'find', 'styleId');
		const replaceAlign = field<HTMLSelectElement>(panel.element, 'replace', 'alignment');
		const replaceStyle = field<HTMLSelectElement>(panel.element, 'replace', 'styleId');
		const replaceAllBtn = panel.element.querySelector<HTMLButtonElement>(
			'button[aria-label="Replace all"]',
		)!;

		// Empty Find text + Heading1 style -> matches "Chapter One"
		findStyle.value = 'Heading1';
		findStyle.dispatchEvent(new Event('change'));
		expect(status.textContent).toContain('1');

		// Empty Replace text + Center alignment + Heading2 style -> keeps text, updates paragraph
		replaceAlign.value = 'center';
		replaceAlign.dispatchEvent(new Event('change'));
		replaceStyle.value = 'Heading2';
		replaceStyle.dispatchEvent(new Event('change'));
		replaceAllBtn.click();

		const saved = docToModel(view.state.doc, model);
		const p1 = saved.blocks[0]!;
		const p2 = saved.blocks[1]!;
		if (p1.type === 'paragraph' && p2.type === 'paragraph') {
			expect(p1.runs.map((r) => r.text).join('')).toBe('Chapter One');
			expect(p1.align).toBe('center');
			expect(p1.style).toBe('Heading2');
			expect(p2.style).toBe('Normal');
			expect(p2.align).toBe('left');
		}

		// Clicking No Formatting clears active criteria and summary
		const clearFind = panel.element.querySelector<HTMLButtonElement>(
			'button[aria-label="Find format: No Formatting"]',
		)!;
		clearFind.click();
		expect(panel.element.querySelector('[data-format-summary="find"]')?.textContent).toBe('');
		expect(status.textContent).toContain('Enter text to search');

		view.destroy();
	});
});
