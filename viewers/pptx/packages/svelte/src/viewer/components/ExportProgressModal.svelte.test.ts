import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import ExportProgressModal from './ExportProgressModal.svelte';

/** The modal renders inside the shared element's open shadow root. */
const inner = (target: HTMLElement, selector: string): HTMLElement | null =>
	target
		.querySelector('pptx-ui-export-progress')
		?.shadowRoot?.querySelector<HTMLElement>(selector) ?? null;

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function open(props: Record<string, unknown>) {
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(ExportProgressModal, {
		target,
		props: { title: 'Export as PDF', progress: 0, oncancel: vi.fn(), ...props } as never,
	});
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	flushSync();
	return target;
}

describe('exportProgressModal', () => {
	it('is the shared progress dialog with a bar and the status line', () => {
		const target = open({ open: true, progress: 62.5, statusMessage: 'Rendering slide 5 of 8...' });
		expect(inner(target, '[role="dialog"]')?.getAttribute('aria-modal')).toBe('true');
		expect(inner(target, '#title')?.textContent).toBe('Export as PDF');
		expect(inner(target, '[role="progressbar"]')?.getAttribute('aria-valuenow')).toBe('63');
		expect(inner(target, '#status')?.textContent).toBe('Rendering slide 5 of 8...');
	});

	it('hands Cancel to the host, and draws nothing while closed', () => {
		const oncancel = vi.fn();
		const shown = open({ open: true, oncancel });
		expect(inner(shown, '#status')?.textContent).toBe('Processing…');
		inner(shown, '[data-testid="pptx-export-progress-cancel"]')!.click();
		expect(oncancel).toHaveBeenCalledOnce();
		cleanup?.();
		const hidden = open({ open: false });
		expect(hidden.querySelector('pptx-ui-export-progress')).not.toBeNull();
		expect(inner(hidden, '[role="dialog"]')).toBeNull();
	});
});
