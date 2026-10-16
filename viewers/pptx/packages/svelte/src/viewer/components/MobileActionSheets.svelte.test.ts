import { registerPptxWebControls } from 'ooxml-ui/pptx';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../editor/editor-state.svelte';
import MobileActionSheets from './MobileActionSheets.svelte';

registerPptxWebControls();

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function open(props: Record<string, unknown> = {}) {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [] }]);
	const onactivechange = vi.fn();
	const target = document.createElement('div');
	const instance = mount(MobileActionSheets, {
		target,
		props: {
			active: null,
			onactivechange,
			editor,
			slides: [{ id: 's1' }, { id: 's2' }],
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map(),
			current: 0,
			onselect: vi.fn(),
			...props,
		},
	});
	flushSync();
	cleanup = () => unmount(instance);
	const buttons = Array.from(
		target.querySelector('pptx-ui-mobile-bar')!.shadowRoot!.querySelectorAll('button'),
	);
	return { buttons, editor, onactivechange, target };
}

describe('mobileActionSheets bottom bar adapter', () => {
	it('opens the tapped sheet, and Insert opens the shared phone Insert sheet', () => {
		const { buttons, onactivechange } = open();
		buttons[0].click();
		expect(onactivechange).toHaveBeenLastCalledWith('slides');
		buttons[1].click();
		expect(onactivechange).toHaveBeenLastCalledWith('insert');
	});

	it('draws the phone Insert sheet and a Text Box tile inserts through the editor, then closes', () => {
		const { editor, onactivechange, target } = open({ active: 'insert' });
		const sheet = target.querySelector('pptx-ui-mobile-insert')!;
		sheet
			.shadowRoot!.querySelector<HTMLButtonElement>('[data-insert-control="insert.text.textBox"]')!
			.click();
		expect(editor.slides[0]?.elements).toHaveLength(1);
		expect(onactivechange).toHaveBeenLastCalledWith(null);
	});

	it('toggles the open sheet closed and reflects it as pressed', () => {
		const { buttons, onactivechange } = open({ active: 'notes' });
		expect(buttons[4].getAttribute('aria-pressed')).toBe('true');
		buttons[4].click();
		expect(onactivechange).toHaveBeenLastCalledWith(null);
	});

	it('disables every slot with no slides', () => {
		const { buttons } = open({ slides: [] });
		expect(buttons.every((button) => button.disabled)).toBeTruthy();
	});
});
