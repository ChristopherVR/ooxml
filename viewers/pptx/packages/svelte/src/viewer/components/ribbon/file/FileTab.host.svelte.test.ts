/**
 * The File backstage covers its viewer host, not the page around it, and Escape pressed inside
 * it closes it while keys pressed elsewhere on the page are left alone.
 */
import { flushSync, mount, unmount } from 'svelte';
import { describe, expect, it, vi } from 'vitest';

import FileTabCloudHarness from './FileTabCloudHarness.svelte';

describe('fileTab host containment', () => {
	it('sits on the viewer host box and closes on Escape pressed inside it', () => {
		const outside = document.createElement('button');
		const host = document.createElement('div');
		host.setAttribute('data-pptx-theme-root', '');
		document.body.append(outside, host);
		vi.spyOn(host, 'getBoundingClientRect').mockReturnValue({
			left: 160,
			top: 40,
			width: 1000,
			height: 700,
		} as DOMRect);
		const onclose = vi.fn();
		const instance = mount(FileTabCloudHarness, {
			target: host,
			props: { cloud: undefined, onclose, oncreatepresentation: vi.fn() },
		});
		flushSync();
		const panel = host.querySelector<HTMLElement>('[role="dialog"]')!;
		expect([
			panel.style.left,
			panel.style.top,
			panel.style.width,
			panel.style.height,
		]).toStrictEqual(['160px', '40px', '1000px', '700px']);
		const escape = (target: Element) =>
			target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		escape(outside);
		expect(onclose).not.toHaveBeenCalled();
		escape(panel.querySelector('button')!);
		expect(onclose).toHaveBeenCalledOnce();
		unmount(instance);
		host.remove();
		outside.remove();
		vi.restoreAllMocks();
	});
});
