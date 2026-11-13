// @vitest-environment happy-dom
/**
 * A stage rebuild that throws (outside any one element, which the stage omits on its own) must
 * leave the last good canvas on screen, not an emptied stage host.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { Translator } from './i18n';
import { createDefaultRegistry } from './render';
import { createRenderController } from './render-controller';
import type { RenderControllerDeps } from './render-controller';
import { createInitialViewerState, createStore } from './state';
import type { ViewerChrome } from './ui';

const slide: PptxSlide = {
	id: 's1',
	rId: 'r1',
	slideNumber: 1,
	elements: [{ id: 'a', type: 'shape', x: 0, y: 0, width: 50, height: 50, shapeType: 'rect' }],
};

describe('vanilla stage rebuild', () => {
	it('keeps the previous stage when the rebuild throws', () => {
		const store = createStore({ ...createInitialViewerState(), slides: [slide], zoom: 1 });
		const stageWrap = document.createElement('div');
		const chrome = {
			stageWrap,
			viewport: document.createElement('div'),
			setEmpty: () => {},
			presentationTouchControls: { update: () => {} },
			notes: { update: () => {} },
		} as unknown as ViewerChrome;
		let fail = false;
		const translator: Translator = ((key: string) => key) as Translator;
		const deps = {
			doc: document,
			store,
			registry: createDefaultRegistry(),
			getChrome: () => chrome,
			getTranslator: () => {
				if (fail) throw new Error('translator unavailable');
				return translator;
			},
		};
		// Every other dependency is a getter or callback this render does not need.
		const renderer = createRenderController(
			new Proxy(deps, {
				get: (target, key) =>
					key in target ? target[key as keyof typeof target] : () => undefined,
			}) as unknown as RenderControllerDeps,
		);
		renderer.renderStage();
		const stage = stageWrap.querySelector('.pptxv-stage');
		expect(stage?.querySelector('[data-element-id="a"]')).not.toBeNull();
		fail = true;
		expect(() => renderer.renderStage()).toThrow('translator unavailable');
		expect(stageWrap.querySelector('.pptxv-stage')).toBe(stage);
	});
});
