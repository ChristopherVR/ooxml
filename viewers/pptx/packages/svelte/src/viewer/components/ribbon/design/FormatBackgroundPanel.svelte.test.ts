import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import FormatBackgroundPanel from './FormatBackgroundPanel.svelte';

/**
 * The docked Format Background panel closes on Escape. It always had an Escape
 * handler, but on the panel itself, and opening it left focus on the ribbon
 * command that opened it, so the key never reached the panel.
 */

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function makeEditor(): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [] }]);
	return editor;
}

describe('formatBackgroundPanel', () => {
	it('takes focus when opened and closes on Escape, handing focus back', () => {
		const opener = document.createElement('button');
		const target = document.createElement('div');
		document.body.append(opener, target);
		const closed = vi.fn();
		const props = $state({
			editor: makeEditor(),
			open: false,
			onclose: () => {
				closed();
				props.open = false;
			},
		});
		const instance = mount(FormatBackgroundPanel, { target, props });
		cleanup = () => {
			unmount(instance);
			target.remove();
			opener.remove();
		};
		opener.focus();

		props.open = true;
		flushSync();
		const panel = target.querySelector<HTMLElement>('[role="dialog"]')!;
		expect(panel.contains(document.activeElement)).toBeTruthy();

		const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
		document.activeElement!.dispatchEvent(escape);
		flushSync();
		expect(closed).toHaveBeenCalledOnce();
		expect(escape.defaultPrevented).toBeTruthy();
		expect(target.querySelector('[role="dialog"]')).toBeNull();
		expect(document.activeElement).toBe(opener);
	});
});
