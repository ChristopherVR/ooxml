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
});
