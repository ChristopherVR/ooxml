/**
 * The chords the inline text editor owns: Ctrl/Cmd+Enter ends the edit
 * (keeping the text, as React and Angular already did), and Ctrl/Cmd+D is
 * cancelled so the browser's bookmark dialog cannot open mid-edit. Both are
 * resolved by the shared `mapInlineEditorChord`.
 */
import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { describe, expect, it, vi } from 'vitest';

import InlineTextEditor from './InlineTextEditor.svelte';

function textElement(): PptxElement {
	return {
		id: 'el-1',
		type: 'text',
		x: 0,
		y: 0,
		width: 200,
		height: 50,
		text: 'TARGET',
		textSegments: [{ text: 'TARGET', style: {} }],
	} as unknown as PptxElement;
}

function open() {
	const host = document.createElement('div');
	document.body.appendChild(host);
	const oncommit = vi.fn();
	const onclose = vi.fn();
	const component = mount(InlineTextEditor, {
		target: host,
		props: {
			element: textElement(),
			box: { x: 0, y: 0, width: 200, height: 50 },
			scale: 1,
			oncommit,
			onclose,
		},
	});
	flushSync();
	const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
	return {
		editor,
		oncommit,
		onclose,
		done: () => {
			unmount(component);
			host.remove();
		},
	};
}

const chord = (key: string): KeyboardEvent =>
	new KeyboardEvent('keydown', { key, ctrlKey: true, bubbles: true, cancelable: true });

describe('inline text editor chords', () => {
	it('ends the edit on Ctrl+Enter, keeping the text', () => {
		const { editor, oncommit, onclose, done } = open();
		editor.textContent = 'CHANGED';
		const event = chord('Enter');
		editor.dispatchEvent(event);
		flushSync();
		expect(event.defaultPrevented).toBeTruthy();
		expect(oncommit).toHaveBeenCalledOnce();
		expect(onclose).toHaveBeenCalledOnce();
		done();
	});

	it('cancels Ctrl+D (the browser bookmark dialog) and stays in the edit', () => {
		const { editor, oncommit, onclose, done } = open();
		const event = chord('d');
		editor.dispatchEvent(event);
		flushSync();
		expect(event.defaultPrevented).toBeTruthy();
		expect(oncommit).not.toHaveBeenCalled();
		expect(onclose).not.toHaveBeenCalled();
		done();
	});

	it('types a tab on Tab and keeps Shift+Tab in the edit', () => {
		const { editor, oncommit, onclose, done } = open();
		editor.focus();
		const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
		editor.dispatchEvent(tab);
		flushSync();
		expect(tab.defaultPrevented).toBe(true);
		expect(editor.textContent).toContain('\t');
		const back = new KeyboardEvent('keydown', {
			key: 'Tab',
			shiftKey: true,
			bubbles: true,
			cancelable: true,
		});
		editor.dispatchEvent(back);
		flushSync();
		expect(back.defaultPrevented).toBe(true);
		expect(oncommit).not.toHaveBeenCalled();
		expect(onclose).not.toHaveBeenCalled();
		done();
	});

	it('demotes a bulleted paragraph on Tab at start and promotes on Shift+Tab', () => {
		const host = document.createElement('div');
		document.body.appendChild(host);
		const oninput = vi.fn();
		const bulletEl = {
			...textElement(),
			text: 'Item',
			textSegments: [{ text: 'Item', style: {}, bulletInfo: { char: '•' }, paragraphLevel: 0 }],
		} as unknown as PptxElement;
		const component = mount(InlineTextEditor, {
			target: host,
			props: {
				element: bulletEl,
				box: { x: 0, y: 0, width: 200, height: 50 },
				scale: 1,
				oncommit: vi.fn(),
				onclose: vi.fn(),
				oninput,
			},
		});
		flushSync();
		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
		const textNode = walker.nextNode() as Text;
		const range = document.createRange();
		range.setStart(textNode, 0);
		range.setEnd(textNode, 0);
		const sel = window.getSelection()!;
		sel.removeAllRanges();
		sel.addRange(range);

		editor.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }),
		);
		flushSync();
		expect(oninput).toHaveBeenLastCalledWith(
			'Item',
			expect.objectContaining({
				textSegments: [expect.objectContaining({ paragraphLevel: 1 })],
			}),
		);

		editor.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true }),
		);
		flushSync();
		expect(oninput).toHaveBeenLastCalledWith(
			'Item',
			expect.objectContaining({
				textSegments: [expect.objectContaining({ paragraphLevel: 0 })],
			}),
		);
		unmount(component);
		host.remove();
	});
});
