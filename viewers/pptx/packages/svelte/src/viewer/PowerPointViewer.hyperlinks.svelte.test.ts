import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushSync, mount, unmount } from 'svelte';
import { expect, test, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.svelte';
import type { PowerPointViewerApi } from './types';

test('offers the host a cancellable read-only text jump without opening a tab', async () => {
	const source = new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/slide-jump-links.pptx')),
	);
	const target = document.createElement('div');
	document.body.append(target);
	const onload = vi.fn();
	const onhyperlinkclick = vi.fn();
	const open = vi.spyOn(window, 'open').mockReturnValue(null);
	const instance = mount(PowerPointViewer, {
		target,
		props: { source, editable: false, onload, onhyperlinkclick },
	});
	try {
		flushSync();
		await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
		flushSync();
		const api = instance as unknown as PowerPointViewerApi;
		const link = target.querySelector<HTMLElement>('[data-pptx-viewport] [data-pptx-hyperlink]')!;
		expect(link.getAttribute('href')).toBeNull();
		onhyperlinkclick.mockReturnValueOnce(false);
		link.click();
		flushSync();
		expect(api.getActiveSlideIndex()).toBe(0);
		link.click();
		flushSync();
		expect(api.getActiveSlideIndex()).toBe(2);
		expect(onhyperlinkclick).toHaveBeenCalledTimes(2);
		expect(open).not.toHaveBeenCalled();
	} finally {
		await unmount(instance);
		target.remove();
		vi.restoreAllMocks();
	}
});
