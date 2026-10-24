/**
 * colour-drag.test-support.ts: the editor the vanilla live colour drag suites
 * drive, wired end to end: one store over a one-slide deck, the real edit
 * actions and history, and the chrome's inspector handlers the sections call.
 *
 * Not a `.test.ts` file on purpose: a shared harness is not a suite.
 */
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { vi } from 'vitest';

import { buildChromeCallbacks } from '../../chrome-callbacks';
import type { ChromeCallbackDeps } from '../../chrome-callbacks';
import { createEditActions } from '../../editor/editor-edit-ops';
import { createEditorOps } from '../../editor/editor-operations';
import { createTranslator } from '../../i18n';
import type { ViewerState } from '../../state';
import { createInitialViewerState, createStore } from '../../state';

/** An editable deck of one slide of `elements`, with `selection` selected, and its inspector handlers. */
export function mountInspectorEditor(elements: PptxElement[], selection: Partial<ViewerState>) {
	const slide = {
		id: 's',
		rId: 'rId1',
		slideNumber: 1,
		notes: '',
		elements,
	} as PptxSlide;
	const store = createStore({
		...createInitialViewerState(),
		slides: [slide],
		currentSlide: 0,
		editable: true,
		...selection,
		presentationProperties: { mruColors: [] },
	});
	const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
	const actions = createEditActions({
		doc: document,
		getTranslator: () => createTranslator(),
		store,
		ops,
		getHandler: () => null,
		setHandler: () => {},
	});
	const deps = new Proxy({ getEditActions: () => actions } as unknown as ChromeCallbackDeps, {
		get: (target, property, receiver) => Reflect.get(target, property, receiver) ?? vi.fn(),
	});
	return { store, ops, handlers: buildChromeCallbacks(deps).inspectorHandlers };
}
