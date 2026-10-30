import type { PptxCoreProperties, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import type { PptxAiBridgePort } from './bridge-factory';
import { createPptxAiBridge, pptxAiDeckTitle } from './bridge-factory';
import { makeSlide, textElement } from './mock-bridge';

interface PortState {
	slides: PptxSlide[];
	active: number;
	canvas: { width: number; height: number };
	core: PptxCoreProperties | undefined;
	commits: string[];
}

function makePort(overrides: Partial<PptxAiBridgePort> = {}): {
	port: PptxAiBridgePort;
	state: PortState;
} {
	const state: PortState = {
		slides: [
			makeSlide(0, [textElement('a', '  Quarterly review  '), textElement('b', 'Body')]),
			makeSlide(1, [textElement('c', 'Two')]),
		],
		active: 0,
		canvas: { width: 960, height: 540 },
		core: { title: 'Deck' },
		commits: [],
	};
	const port: PptxAiBridgePort = {
		getSlides: () => state.slides,
		getActiveSlideIndex: () => state.active,
		getCanvasSize: () => state.canvas,
		getTheme: () => undefined,
		getHandler: () => null,
		getSections: () => [],
		getPresentationProperties: () => ({}),
		getCustomProperties: () => [],
		getCoreProperties: () => state.core,
		getAppProperties: () => undefined,
		getViewProperties: () => undefined,
		getTableStyleMap: () => undefined,
		getTableStylesDefaultId: () => undefined,
		getTagCollections: () => [],
		goToSlide: (index) => {
			state.active = index;
		},
		selectElements: vi.fn(),
		updateSlides: (update, label) => {
			state.slides = update(state.slides);
			state.commits.push(label);
		},
		applyTheme: vi.fn(),
		setCanvasSize: (size) => {
			state.canvas = size;
		},
		setSections: vi.fn(),
		setPresentationProperties: vi.fn(),
		setDocumentProperties: (core) => {
			state.core = core;
		},
		setViewProperties: vi.fn(),
		setTableStyleMap: vi.fn(),
		setTableStylesDefaultId: vi.fn(),
		setTagCollections: vi.fn(),
		...overrides,
	};
	return { port, state };
}

describe('createPptxAiBridge', () => {
	it('reports deck metadata with the file name, else the trimmed first text', () => {
		const { port } = makePort();
		expect(createPptxAiBridge(port).getDeckMeta()).toEqual({
			slideCount: 2,
			activeSlideIndex: 0,
			title: 'Quarterly review',
			width: 960,
			height: 540,
		});
		const named = createPptxAiBridge({ ...port, getFileName: () => 'Plan.pptx' });
		expect(named.getDeckMeta().title).toBe('Plan.pptx');
	});

	it('commits an element update as one labelled slides edit without mutating the live deck', () => {
		const { port, state } = makePort();
		const before = state.slides;
		createPptxAiBridge(port).updateElement(0, 'b', { x: 99 });
		expect(state.commits).toEqual(['Update element b']);
		expect(state.slides[0]?.elements[1]?.x).toBe(99);
		expect(before[0]?.elements[1]?.x).toBe(40);
	});

	it('applies two updates in one tick on top of each other', () => {
		const { port, state } = makePort();
		const bridge = createPptxAiBridge(port);
		bridge.updateElement(0, 'a', { x: 1 });
		bridge.updateElement(0, 'b', { x: 2 });
		expect(state.slides[0]?.elements.map((el) => el.x)).toEqual([1, 2]);
	});

	it('fans only the changed deck fields out, then calls afterDeckUpdate once', () => {
		const afterDeckUpdate = vi.fn();
		const onCorePropertiesChanged = vi.fn();
		const { port, state } = makePort({ afterDeckUpdate, onCorePropertiesChanged });
		const bridge = createPptxAiBridge(port);
		bridge.applyDeckData?.(
			(data) => ({ ...data, width: 1280, height: 720, coreProperties: { title: 'New' } }),
			'Resize',
		);
		expect(state.canvas).toEqual({ width: 1280, height: 720 });
		expect(state.core).toEqual({ title: 'New' });
		expect(state.commits).toEqual([]);
		expect(onCorePropertiesChanged).toHaveBeenCalledWith({ title: 'New' });
		expect(port.setSections).not.toHaveBeenCalled();
		expect(afterDeckUpdate).toHaveBeenCalledOnce();

		bridge.applyDeckData?.((data) => data, 'No-op');
		expect(afterDeckUpdate).toHaveBeenCalledOnce();
	});

	it('commits changed custom shows only when the port can set them', () => {
		const setCustomShows = vi.fn();
		const shows = [{ id: '0', name: 'Short', slideRIds: ['rId2'] }];
		const { port } = makePort({ getCustomShows: () => [], setCustomShows });
		createPptxAiBridge(port).applyDeckData?.((data) => ({ ...data, customShows: shows }), 'Shows');
		expect(setCustomShows).toHaveBeenCalledWith(shows);

		const { port: readOnly } = makePort();
		expect(() =>
			createPptxAiBridge(readOnly).applyDeckData?.(
				(data) => ({ ...data, customShows: shows }),
				'Shows',
			),
		).not.toThrow();
	});

	it('exposes focus and notify only when the port provides them', () => {
		const { port } = makePort();
		expect(createPptxAiBridge(port).getFocusedTargets).toBeUndefined();
		const focused = createPptxAiBridge({
			...port,
			getFocusedTargets: () => [{ kind: 'slide', slideIndex: 1 }],
		});
		expect(focused.getFocusedTargets?.()).toEqual([{ kind: 'slide', slideIndex: 1 }]);
	});
});

describe('pptxAiDeckTitle', () => {
	it('returns undefined for a deck with no text on the first slide', () => {
		expect(pptxAiDeckTitle([makeSlide(0, [])])).toBeUndefined();
	});
});
