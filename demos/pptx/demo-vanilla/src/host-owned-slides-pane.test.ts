import {
	createCollaborationShell,
	createInitialViewerState,
	createStore,
} from 'pptx-vanilla-viewer';
import type { PptxSlide } from 'pptx-vanilla-viewer';
import { readSlidesFromYDoc, reconcileSlidesInYDoc } from 'pptx-viewer-shared';
import { describe, expect, it } from 'vitest';
import { Awareness } from 'y-protocols/awareness';
import * as Y from 'yjs';

import { mountHostOwnedSlidesPane } from './host-owned-slides-pane';

const slides = (count: number): PptxSlide[] =>
	Array.from({ length: count }, (_, index) => ({
		id: `slide-${index}`,
		rId: `rId${index}`,
		slideNumber: index + 1,
		elements: [],
	}));

async function setup() {
	const doc = new Y.Doc();
	const awareness = new Awareness(doc);
	const store = createStore(createInitialViewerState());
	let allowed = true;
	const shell = createCollaborationShell({
		document,
		store,
		getHandler: () => null,
		getCanEdit: () => allowed,
	});
	reconcileSlidesInYDoc(slides(3), doc, {
		createMap: () => new Y.Map(),
		createArray: () => new Y.Array(),
		createText: () => new Y.Text(),
	});
	await shell.setConfig({
		roomId: 'custom-shell-rail',
		serverUrl: '',
		userName: 'Host',
		externalSession: {
			doc,
			awareness,
			getSnapshot: () => ({ status: 'connected', synced: true }),
			subscribe: () => () => {},
		},
	});
	const root = document.createElement('div');
	root.dataset.hostCustomShell = '';
	document.body.append(root);
	let inlineOpen = false;
	const pane = mountHostOwnedSlidesPane(root, {
		store,
		getLoaded: () => undefined,
		commitInline: () => !inlineOpen,
	});
	return {
		doc,
		store,
		root,
		rows: () =>
			Array.from(root.querySelectorAll<HTMLButtonElement>('[data-pptx-chrome="slide-row"]')),
		footerButton: () =>
			root.querySelector<HTMLButtonElement>('[data-pptx-chrome="slide-footer"] button'),
		holdInline(open: boolean) {
			inlineOpen = open;
		},
		disable() {
			allowed = false;
			shell.refresh();
		},
		dispose() {
			pane.destroy();
			root.remove();
			shell.destroy();
			awareness.destroy();
			doc.destroy();
		},
	};
}

describe('custom-shell slide rail', () => {
	it('follows the shared deck, moves the canvas slide and publishes a footer add', async () => {
		const h = await setup();
		try {
			expect(h.store.get().slides).toHaveLength(3);
			expect(h.rows()).toHaveLength(3);
			expect(h.rows()[0].getAttribute('aria-current')).toBe('true');

			h.rows()[2].click();
			expect(h.store.get().currentSlide).toBe(2);
			expect(h.rows()[2].getAttribute('aria-current')).toBe('true');

			h.footerButton()!.click();
			expect(h.store.get().slides).toHaveLength(4);
			expect(h.store.get().currentSlide).toBe(3);
			expect(h.store.get().slides.map((slide) => slide.slideNumber)).toStrictEqual([1, 2, 3, 4]);
			expect(h.rows()).toHaveLength(4);
			expect(h.rows()[3].getAttribute('aria-current')).toBe('true');
			// The peer-visible document has the new slide too.
			expect(readSlidesFromYDoc(h.doc).map((slide) => slide.id)).toStrictEqual(
				h.store.get().slides.map((slide) => slide.id),
			);
		} finally {
			h.dispose();
		}
	});

	it('keeps an open inline edit and a read-only session from changing slides', async () => {
		const h = await setup();
		try {
			h.holdInline(true);
			h.rows()[1].click();
			h.footerButton()!.click();
			expect(h.store.get().currentSlide).toBe(0);
			expect(h.store.get().slides).toHaveLength(3);

			h.holdInline(false);
			h.disable();
			expect(h.store.get().editable).toBe(false);
			expect(h.footerButton()?.closest<HTMLElement>('[hidden]')).not.toBeNull();
			h.footerButton()?.click();
			expect(h.store.get().slides).toHaveLength(3);
			h.rows()[1].click();
			expect(h.store.get().currentSlide).toBe(1);
		} finally {
			h.dispose();
		}
	});
});
