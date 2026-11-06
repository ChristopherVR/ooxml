/**
 * A slide is painted several times on one page (the editor canvas and its
 * thumbnails). The ids of its SVG defs derive from the element id, so the copies
 * shared them and `url(#id)` resolved to the first copy, which may sit in a
 * hidden subtree. Each copy now owns its ids and references resolve within it.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { afterEach, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

const deck = new Uint8Array(
	readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/connector-arrows.pptx')),
);
let viewer: PptxViewerInstance | undefined;
afterEach(() => {
	viewer?.destroy();
	document.body.replaceChildren();
});

it('gives every painted copy of a slide its own def ids, resolved within the copy', async () => {
	const container = document.createElement('div');
	document.body.append(container);
	viewer = createPptxViewer(container, { readOnly: true });
	await viewer.loadFile(deck);
	await vi.waitFor(() => {
		expect(container.querySelectorAll('marker').length).toBeGreaterThan(1);
	});
	await new Promise((done) => setTimeout(done, 20));
	const ids = [...container.querySelectorAll('[id]')].map((n) => n.id);
	expect(new Set(ids).size).toBe(ids.length);
	const refs = [...container.querySelectorAll('[marker-end],[marker-start]')];
	expect(refs.length).toBeGreaterThan(0);
	for (const ref of refs) {
		const token = /#([^)]+)\)/u.exec(
			ref.getAttribute('marker-end') ?? ref.getAttribute('marker-start') ?? '',
		)?.[1];
		const owner = ref.closest('[data-element-id]');
		expect(owner?.querySelector(`[id="${token}"]`)).not.toBeNull();
	}
});
