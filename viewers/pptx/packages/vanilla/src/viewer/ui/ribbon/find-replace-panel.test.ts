import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createFindReplaceActions } from '../../editor/editor-find-replace-actions';
import type { EditorOps } from '../../editor/editor-operations';
import { createTranslator } from '../../i18n';
import { createInitialViewerState, createStore } from '../../state';
import { createFindReplacePanel } from './find-replace-panel';

function text(id: string, value: string) {
	return {
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 20,
		text: value,
		textSegments: [{ text: value, style: {} }],
	};
}

/** Three matches of "quarter" on two slides. */
const deck = (): PptxSlide[] =>
	[
		{ id: 's1', rId: '', slideNumber: 1, elements: [text('a1', 'Quarterly growth')] },
		{
			id: 's2',
			rId: '',
			slideNumber: 2,
			elements: [text('b1', 'Quarterly Growth'), text('b2', 'Active users by quarter')],
		},
	] as PptxSlide[];

function panel() {
	const store = createStore({ ...createInitialViewerState(), editable: true, slides: deck() });
	const select = vi.fn();
	const ops = {
		select,
		pushHistory: vi.fn(),
		commitChange: vi.fn(),
	} as unknown as EditorOps;
	const t = createTranslator();
	const view = createFindReplacePanel(document, t, createFindReplaceActions({ store, ops }));
	document.body.append(view.el);
	view.setOpen(true);
	const el = view.el;
	const query = el.querySelector<HTMLInputElement>('input[aria-label="Search text"]')!;
	const button = (label: string) =>
		el.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!;
	return {
		view,
		store,
		select,
		status: () => el.querySelector('.pptxv-findreplace-status')!.textContent,
		type: (value: string) => {
			query.value = value;
			query.dispatchEvent(new Event('input', { bubbles: true }));
		},
		key: (key: string, shiftKey = false) =>
			query.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey, bubbles: true })),
		replacement: el.querySelector<HTMLInputElement>('input[aria-label="Replacement text"]')!,
		matchCase: () => {
			const box = el.querySelector<HTMLInputElement>('[aria-label="Toggle match case"]')!;
			box.checked = !box.checked;
			box.dispatchEvent(new Event('change', { bubbles: true }));
		},
		next: () => button('Next match').click(),
		previous: () => button('Previous match').click(),
		replace: () => button('Replace current match').click(),
	};
}

afterEach(() => {
	document.body.replaceChildren();
});

describe('createFindReplacePanel', () => {
	it('next and previous move through the matches with a real n of N, wrapping around', () => {
		const p = panel();
		p.type('quarter');
		expect(p.status()).toBe('1 of 3');
		p.next();
		expect(p.status()).toBe('2 of 3');
		expect(p.store.get().currentSlide).toBe(1);
		expect(p.select).toHaveBeenLastCalledWith('b1');
		p.previous();
		p.previous();
		expect(p.status()).toBe('3 of 3');
	});

	it('enter goes to the next match and shift+enter to the previous one', () => {
		const p = panel();
		p.type('quarter');
		p.key('Enter');
		expect(p.status()).toBe('2 of 3');
		p.key('Enter', true);
		p.key('Enter', true);
		expect(p.status()).toBe('3 of 3');
	});

	it('replace changes the current match only', () => {
		const p = panel();
		p.type('quarter');
		p.next();
		p.replacement.value = 'period';
		p.replace();
		const texts = p.store
			.get()
			.slides.flatMap((s) => s.elements.map((e) => (e as { text: string }).text));
		expect(texts).toStrictEqual(['Quarterly growth', 'periodly Growth', 'Active users by quarter']);
		expect(p.status()).toBe('2 of 2');
	});

	it('match case searches again', () => {
		const p = panel();
		p.type('Quarter');
		expect(p.status()).toBe('1 of 3');
		p.matchCase();
		expect(p.status()).toBe('1 of 2');
	});

	it('closing drops the count; reopening searches the kept query from the first match', () => {
		const p = panel();
		p.type('quarter');
		p.next();
		p.view.setOpen(false);
		expect(p.status()).toBe('');
		p.view.setOpen(true);
		expect(p.status()).toBe('1 of 3');
	});
});
