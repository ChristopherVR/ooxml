// @vitest-own-worker: mounts full viewers; see region-customization.test-support.ts.
import { afterEach, describe, expect, it, vi } from 'vitest';

import { destroyMounted, mount } from './region-customization.test-support';
import { ownOverlay } from './ui/overlay-owner';

afterEach(destroyMounted);

/** The dialog openers the class has but the public instance type leaves out. */
type Openers = Record<
	| 'openSettings'
	| 'openPrintDialog'
	| 'openHeaderFooter'
	| 'openDocumentProperties'
	| 'openFontEmbedding'
	| 'openPasswordProtection',
	() => void
>;

const OVERLAYS =
	'pptx-ui-options-dialog, .pptxv-parity-dialog, .pptxv-parity-backdrop, .pptxv-props-overlay, .pptxv-info-overlay';

describe('dialogs the viewer mounts on document.body', () => {
	it('close with the viewer instead of outliving it', () => {
		const { viewer } = mount({ editable: true });
		const open = viewer as unknown as Openers;
		open.openSettings();
		open.openPrintDialog();
		open.openHeaderFooter();
		open.openDocumentProperties();
		open.openFontEmbedding();
		open.openPasswordProtection();
		expect(document.body.querySelectorAll(OVERLAYS).length).toBeGreaterThanOrEqual(6);
		viewer.destroy();
		expect(document.body.querySelectorAll(OVERLAYS)).toHaveLength(0);
	});

	it('leave the other viewer alone, and a dialog closed by the user is released', () => {
		const first = mount({ editable: true });
		const second = mount({ editable: true });
		(second.viewer as unknown as Openers).openPrintDialog();
		first.viewer.destroy();
		expect(document.body.querySelector('.pptxv-parity-dialog')).not.toBeNull();
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(document.body.querySelector('.pptxv-parity-dialog')).toBeNull();
		second.viewer.destroy();
	});
});

describe('ownOverlay', () => {
	it('unregisters on close and closes once on the owner going away', () => {
		const teardowns = new Set<() => void>();
		const owner = {
			own(close: () => void) {
				teardowns.add(close);
				return () => teardowns.delete(close);
			},
		};
		const firstClose = vi.fn();
		const closeFirst = ownOverlay(owner, firstClose);
		const secondClose = vi.fn();
		ownOverlay(owner, secondClose);
		closeFirst();
		expect(teardowns.size).toBe(1);
		for (const teardown of [...teardowns]) teardown();
		expect(firstClose).toHaveBeenCalledOnce();
		expect(secondClose).toHaveBeenCalledOnce();
		expect(teardowns.size).toBe(0);
	});

	it('returns the close unchanged without an owner', () => {
		const close = vi.fn();
		expect(ownOverlay(undefined, close)).toBe(close);
	});
});
