/**
 * read-only-password-timing.test.ts: a modify-password check that finishes late. The check
 * (`checkModifyPassword`) is replaced by one the test resolves by hand, so a new load or Cancel
 * can be made to land before the result, which is otherwise a race against the real digest.
 */
import type { ModifyPasswordCheckResult } from 'ooxml-ui/pptx';
import { createModifyVerifier, PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

const pending: Array<(result: ModifyPasswordCheckResult) => void> = [];
// A string, not `import(...)`: the module is outside this package's `rootDir`, so the type
// checker must not follow it. It is the module the shared session runs every check through.
// oxlint-disable-next-line vitest/prefer-import-in-mock
vi.mock('../../../../../../src/ui/src/pptx/render/modify-password-check', () => ({
	checkModifyPassword: () =>
		new Promise<ModifyPasswordCheckResult>((resolve) => {
			pending.push(resolve);
		}),
}));

let deck: Uint8Array;
beforeAll(async () => {
	const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	try {
		deck = await handler.save(data.slides, { modifyVerifier: verifier });
	} finally {
		handler.dispose();
	}
}, 30_000);

let active: PptxViewerInstance[] = [];
afterEach(() => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	pending.length = 0;
	document.body.replaceChildren();
});

function inner<T extends HTMLElement = HTMLElement>(container: HTMLElement, id: string): T {
	const banner = container.querySelector('pptx-ui-read-only-banner');
	return banner!.shadowRoot!.querySelector<T>(`[data-testid="${id}"]`)!;
}

/** Load the deck, open the password prompt and submit a password whose check stays pending. */
async function submitPending(): Promise<{ container: HTMLElement; viewer: PptxViewerInstance }> {
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, { editable: true });
	active.push(viewer);
	await viewer.loadFile(deck);
	inner<HTMLButtonElement>(container, 'pptx-readonly-edit-anyway').click();
	inner<HTMLInputElement>(container, 'pptx-readonly-password-input').value = 'right-password';
	inner(container, 'pptx-readonly-password-form').dispatchEvent(
		new Event('submit', { cancelable: true }),
	);
	expect(pending).toHaveLength(1);
	return { container, viewer };
}

async function settle(): Promise<void> {
	for (let i = 0; i < 5; i++) {
		await new Promise((resolve) => {
			setTimeout(resolve, 0);
		});
	}
}

const isEditable = (container: HTMLElement) =>
	container.querySelector('.pptxv')?.classList.contains('pptxv-editable');

describe('vanilla read-only password check timing', () => {
	it('ignores a check that finishes after the next deck loads', async () => {
		const { container, viewer } = await submitPending();
		await viewer.loadFile(deck);
		pending[0]!({ ok: true });
		await settle();

		expect(isEditable(container)).toBeFalsy();
		expect(
			container.querySelector<HTMLElement>('[data-testid="pptx-readonly-banner"]')?.hidden,
		).toBeFalsy();
		expect(inner(container, 'pptx-readonly-password-form').hidden).toBeTruthy();
	});

	it('ignores a check that finishes after Cancel', async () => {
		const { container } = await submitPending();
		inner<HTMLButtonElement>(container, 'pptx-readonly-password-cancel').click();
		pending[0]!({ ok: true });
		await settle();

		expect(isEditable(container)).toBeFalsy();
		expect(inner(container, 'pptx-readonly-password-form').hidden).toBeTruthy();
		expect(inner(container, 'pptx-readonly-password-error').hidden).toBeTruthy();
	});
});
