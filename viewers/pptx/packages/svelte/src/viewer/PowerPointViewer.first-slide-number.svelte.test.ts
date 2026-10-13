/**
 * `p:presentation/@firstSlideNum="45"`: PowerPoint numbers slide `i` as
 * 45 + i, so the `slidenum` field on slide 1 reads 45. The viewer showed the
 * slide's position (1) instead.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushSync, mount, unmount } from 'svelte';
import { expect, test, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.svelte';

test('numbers slide-number fields from the deck start number', async () => {
	const source = new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx')),
	);
	const target = document.createElement('div');
	document.body.append(target);
	const onload = vi.fn();
	const instance = mount(PowerPointViewer, { target, props: { source, editable: false, onload } });
	try {
		flushSync();
		await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
		flushSync();
		const stage = target.querySelector('[data-pptx-viewport]')!.textContent;
		expect(stage).toContain('Slide 45');
		expect(stage).not.toContain('Slide 1');
	} finally {
		await unmount(instance);
		target.remove();
	}
});
