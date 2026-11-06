// @vitest-environment happy-dom
/**
 * A slide is painted several times on one page (the editor canvas and its
 * thumbnails). The ids of its SVG defs derive from the element id, so the copies
 * shared them and `url(#id)` resolved to the first copy, which may sit in a
 * hidden subtree. Each copy now owns its ids and references resolve within it.
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
	const hadActEnvironment = Object.hasOwn(globalThis, 'IS_REACT_ACT_ENVIRONMENT');
	const previousActEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	const container = document.createElement('div');
	document.body.append(container);
	const root = createRoot(container);
	const ref = createRef<Handle>();
	try {
		const content = new Uint8Array(
			readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/connector-arrows.pptx')),
		);
		await act(async () => {
			root.render(<PowerPointViewer ref={ref} content={content} canEdit={false} />);
		});
		await vi.waitFor(
			async () => {
				await act(async () => {
					await new Promise((done) => {
						setTimeout(done, 20);
					});
				});
				expect(container.querySelectorAll('marker').length).toBeGreaterThan(1);
			},
			{ timeout: 15000 },
		);
		// The stage, a morph overlay and a mirror paint the same element again: emulate one.
		const element = container.querySelector('marker')!.closest('[data-element-id]')!;
		const second = element.cloneNode(true) as HTMLElement;
		second.style.display = 'none';
		element.parentElement!.prepend(second);
		await act(async () => {
			await new Promise((done) => {
				setTimeout(done, 20);
			});
		});
		expectOwnDefIds(container);
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
