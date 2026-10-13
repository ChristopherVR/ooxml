// @vitest-environment happy-dom
/**
 * `p:presentation/@firstSlideNum="45"`: PowerPoint numbers slide `i` as
 * 45 + i, so the `slidenum` field on slide 1 reads 45. The viewer showed the
 * slide's position (1) on the canvas and in the thumbnail rail.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, test, vi } from 'vitest';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));
const { PowerPointViewer } = await import('../index');
type Handle = import('../index').PowerPointViewerHandle;

test('numbers slide-number fields from the deck start number', async () => {
	const hadActEnvironment = Object.hasOwn(globalThis, 'IS_REACT_ACT_ENVIRONMENT');
	const previousActEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	const container = document.createElement('div');
	document.body.append(container);
	const root = createRoot(container);
	const ref = createRef<Handle>();
	try {
		const content = new Uint8Array(
			readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx')),
		);
		await act(async () => {
			root.render(<PowerPointViewer ref={ref} content={content} canEdit={false} />);
		});
		await vi.waitFor(
			async () => {
				await act(async () => {
					await new Promise((done) => {
						setTimeout(done, 10);
					});
				});
				expect(ref.current?.getSlideCount()).toBe(3);
			},
			{ timeout: 15000 },
		);
		const viewport = container.querySelector('[data-pptx-viewport]')!;
		expect(viewport.textContent).toContain('Slide 45');
		expect(viewport.textContent).not.toContain('Slide 1');
	} finally {
		await act(async () => root.unmount());
		container.remove();
		if (hadActEnvironment) {
			globalThis.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment;
		} else {
			delete globalThis.IS_REACT_ACT_ENVIRONMENT;
		}
	}
});
