/**
 * A slide is painted several times on one page (the editor canvas and its
 * thumbnails). The ids of its SVG defs derive from the element id, so the copies
 * shared them and `url(#id)` resolved to the first copy, which may sit in a
 * hidden subtree. Each copy now owns its ids and references resolve within it.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushSync, mount, unmount } from 'svelte';
import { expect, test, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.svelte';

function expectOwnDefIds(root: Element): void {
	const ids = [...root.querySelectorAll('[id]')].map((n) => n.id);
	expect(new Set(ids).size).toBe(ids.length);
	const refs = [...root.querySelectorAll('[marker-end],[marker-start]')];
	expect(refs.length).toBeGreaterThan(0);
	for (const ref of refs) {
		const token = /#([^)]+)\)/u.exec(
			ref.getAttribute('marker-end') ?? ref.getAttribute('marker-start') ?? '',
		)?.[1];
		expect(ref.closest('[data-element-id]')?.querySelector(`[id="${token}"]`)).not.toBeNull();
	}
}

test('gives every painted copy of a slide its own def ids', async () => {
	const source = new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/connector-arrows.pptx')),
	);
	const target = document.createElement('div');
	document.body.append(target);
	const onload = vi.fn();
	const instance = mount(PowerPointViewer, { target, props: { source, editable: false, onload } });
	try {
		flushSync();
		await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
		await vi.waitFor(
			async () => {
				flushSync();
				await new Promise((done) => setTimeout(done, 20));
				expect(target.querySelectorAll('marker').length).toBeGreaterThan(1);
			},
			{ timeout: 15000 },
		);
		expectOwnDefIds(target);
	} finally {
		await unmount(instance);
		target.remove();
	}
});
