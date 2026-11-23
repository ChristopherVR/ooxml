import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import RecordTab from './RecordTab.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function mountTab(): {
	target: HTMLElement;
	onfrombeginning: ReturnType<typeof vi.fn>;
	onfromcurrent: ReturnType<typeof vi.fn>;
} {
	const target = document.createElement('div');
	const onfrombeginning = vi.fn();
	const onfromcurrent = vi.fn();
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	const instance = mount(RecordTab, {
		target,
		props: { editor, canvasSize: { width: 960, height: 540 }, onfrombeginning, onfromcurrent },
	});
	flushSync();
	cleanup = () => unmount(instance);
	return { target, onfrombeginning, onfromcurrent };
}

/** Find a control by the accessible name the ribbon-inventory e2e spec reads. */
function byName(target: HTMLElement, name: string): HTMLButtonElement | undefined {
	return [...target.querySelectorAll('pptx-ui-ribbon-command')]
		.map((host) => host.shadowRoot!.querySelector<HTMLButtonElement>('button')!)
		.find((button) => button.textContent?.trim() === name);
}

describe('recordTab', () => {
	it('starts recording from the beginning or current slide', () => {
		const { target, onfrombeginning, onfromcurrent } = mountTab();

		byName(target, 'From Beginning')?.click();
		byName(target, 'From Current Slide')?.click();

		expect(onfrombeginning).toHaveBeenCalledOnce();
		expect(onfromcurrent).toHaveBeenCalledOnce();
	});

	it('disables editing commands in a read-only presentation', () => {
		const { target } = mountTab();

		expect(target.querySelectorAll('pptx-ui-ribbon-section > pptx-ui-ribbon-group')).toHaveLength(
			4,
		);

		for (const name of ['Cameo', 'Clear timings', 'Reset to Cameo', 'Learn More']) {
			const button = byName(target, name);
			expect(button, `${name} is missing from the Record tab`).toBeDefined();
			expect(button?.disabled, `${name} should be disabled`).toBeTruthy();
		}
	});
});
