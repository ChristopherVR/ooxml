import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

const deck = new Uint8Array(
	readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/slide-jump-links.pptx')),
);
let viewer: PptxViewerInstance | undefined;
afterEach(() => {
	viewer?.destroy();
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

describe('vanilla read-only hyperlinks', () => {
	it.each(['A. Text link to slide 3', 'B. Shape link to slide 3'])(
		'navigates %s and offers the host a cancellable callback',
		async (text) => {
			const container = document.createElement('div');
			document.body.append(container);
			const onHyperlinkClick = vi.fn();
			const open = vi.spyOn(window, 'open').mockReturnValue(null);
			viewer = createPptxViewer(container, { readOnly: true, onHyperlinkClick });
			await viewer.loadFile(deck);
			const click = () => {
				const stage = container.querySelector('[data-pptx-viewport]')!;
				const selector = text.startsWith('A.') ? '[data-pptx-hyperlink]' : '[data-element-id]';
				const link = [...stage.querySelectorAll<HTMLElement>(selector)].find((node) =>
					node.textContent?.includes(text),
				)!;
				link.click();
			};
			onHyperlinkClick.mockReturnValueOnce(false);
			click();
			expect(viewer.getActiveSlideIndex()).toBe(0);
			click();
			expect(viewer.getActiveSlideIndex()).toBe(2);
			expect(onHyperlinkClick).toHaveBeenCalledTimes(2);
			expect(onHyperlinkClick).toHaveBeenLastCalledWith(
				expect.objectContaining({ action: 'ppaction://hlinksldjump', targetSlideIndex: 2 }),
			);
			expect(open).not.toHaveBeenCalled();
		},
	);
});
