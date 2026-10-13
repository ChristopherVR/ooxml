/**
 * `p:presentation/@firstSlideNum="45"`: PowerPoint numbers slide `i` as
 * 45 + i, so the `slidenum` field on slide 1 reads 45. The viewer showed the
 * slide's position (1) instead.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { afterEach, expect, it } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

const deck = new Uint8Array(
	readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx')),
);
let viewer: PptxViewerInstance | undefined;
afterEach(() => {
	viewer?.destroy();
	document.body.replaceChildren();
});

it('numbers slide-number fields from the deck start number', async () => {
	const container = document.createElement('div');
	document.body.append(container);
	viewer = createPptxViewer(container, { readOnly: true });
	await viewer.loadFile(deck);
	const stage = container.querySelector('[data-pptx-viewport]')!;
	expect(stage.textContent).toContain('Slide 45');
	expect(stage.textContent).not.toContain('Slide 1');
});
