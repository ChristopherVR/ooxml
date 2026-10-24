import { PRESET_THEMES } from 'ooxml-ui/pptx';
import type {
	PptxElement,
	PptxHandler,
	PptxSlide,
	PptxSlideMaster,
	PptxThemeColorScheme,
} from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import { createDeckActions } from './editor-deck-actions';

function makeSlide(id: string): PptxSlide {
	return { id, rId: `rId-${id}`, slideNumber: 1, elements: [] } as unknown as PptxSlide;
}

function setup(options: { handler?: PptxHandler | null; editable?: boolean } = {}) {
	const store = createStore({
		...createInitialViewerState(),
		editable: options.editable ?? true,
		slides: [makeSlide('s1')],
		slideMasters: [
			{ path: 'ppt/slideMasters/slideMaster1.xml' },
			{ path: 'ppt/slideMasters/slideMaster2.xml' },
		] as unknown as PptxSlideMaster[],
	});
	const ops = {
		pushHistory: vi.fn(),
		commitChange: vi.fn(() => store.set({ dirty: true })),
		updatePresentationProperties: vi.fn(),
	};
	const actions = createDeckActions({
		store,
		ops,
		getHandler: () => options.handler ?? null,
	});
	return { store, ops, actions };
}

describe('editor deck actions (no-selection inspector)', () => {
	it('persists a staged theme before recording one model/history change', async () => {
		const applyTheme = vi.fn().mockResolvedValue(undefined);
		const { actions, ops, store } = setup({ handler: { applyTheme } as unknown as PptxHandler });
		const edit = {
			name: 'Saved edit',
			colorScheme: { ...PRESET_THEMES[0].colorScheme, accent1: '#123456' },
			fontScheme: {},
		};
		await actions.applyThemeEdit(edit);
		expect(applyTheme).toHaveBeenCalledExactlyOnceWith(
			edit.colorScheme,
			edit.fontScheme,
			edit.name,
		);
		expect(ops.pushHistory).toHaveBeenCalledOnce();
		expect(ops.commitChange).toHaveBeenCalledOnce();
		expect(store.get().themeName).toBe(edit.name);
	});

	it('updates the canvas size in the store and marks the deck dirty', () => {
		const { store, actions } = setup();
		actions.updateCanvasSize({ width: 1280.4, height: 720.2 });
		expect(store.get().canvasSize).toStrictEqual({ width: 1280, height: 720 });
		expect(store.get().dirty).toBeTruthy();
	});

	it('ignores canvas-size edits when not editable or not finite', () => {
		const { store, actions } = setup({ editable: false });
		const before = store.get().canvasSize;
		actions.updateCanvasSize({ width: 1280, height: 720 });
		expect(store.get().canvasSize).toBe(before);
		expect(store.get().dirty).toBeFalsy();

		const editable = setup();
		editable.actions.updateCanvasSize({ width: Number.NaN, height: 720 });
		expect(editable.store.get().dirty).toBeFalsy();
	});

	// A theme part by path is applied in place now (`editor-theme-by-path.ts`): the shared step reads
	// the part's schemes and re-points each target master through `setSlideMasterTheme`, and the
	// store takes the result as one undo step. Re-theming the slides is covered with a real deck in
	// `theme-undo.test.ts`.
	function themePartHandler() {
		const schemes = {
			name: 'Facet',
			colorScheme: {
				dk1: '#000000',
				lt1: '#FFFFFF',
				dk2: '#2C3C43',
				lt2: '#EBEBEB',
				accent1: '#90C226',
				accent2: '#54A021',
				accent3: '#E6B91E',
				accent4: '#E76618',
				accent5: '#C42F1A',
				accent6: '#918655',
				hlink: '#99CA3C',
				folHlink: '#B9D181',
			},
			fontScheme: { majorFont: { latin: 'Trebuchet MS' }, minorFont: { latin: 'Trebuchet MS' } },
		};
		const getThemePartSchemes = vi.fn().mockResolvedValue(schemes);
		const setSlideMasterTheme = vi.fn().mockResolvedValue(undefined);
		// The switch resolves the moved slides' style references; these slides have none.
		const reResolveThemeStyleReferences = (
			slides: readonly PptxSlide[],
			templateElementsBySlideId: Readonly<Record<string, PptxElement[]>>,
		) => ({ slides: [...slides], templateElementsBySlideId: { ...templateElementsBySlideId } });
		const handler = {
			getThemePartSchemes,
			setSlideMasterTheme,
			reResolveThemeStyleReferences,
		} as unknown as PptxHandler;
		return { handler, setSlideMasterTheme };
	}

	it('applies a theme by path to the first master only, as one undo step', async () => {
		const { handler, setSlideMasterTheme } = themePartHandler();
		const { store, ops, actions } = setup({ handler });

		await actions.applyThemeByPath('ppt/theme/theme2.xml', false);

		expect(setSlideMasterTheme.mock.calls).toStrictEqual([
			['ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme2.xml'],
		]);
		const masters = store.get().slideMasters as unknown as Array<{ themePath?: string }>;
		expect(masters[0].themePath).toBe('ppt/theme/theme2.xml');
		expect(masters[1].themePath).toBeUndefined();
		expect(ops.pushHistory).toHaveBeenCalledOnce();
		expect(store.get().dirty).toBeTruthy();
	});

	it('applies a theme by path to all masters, as one undo step', async () => {
		const { handler, setSlideMasterTheme } = themePartHandler();
		const { store, ops, actions } = setup({ handler });

		await actions.applyThemeByPath('ppt/theme/theme1.xml', true);

		expect(setSlideMasterTheme.mock.calls).toStrictEqual([
			['ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme1.xml'],
			['ppt/slideMasters/slideMaster2.xml', 'ppt/theme/theme1.xml'],
		]);
		const masters = store.get().slideMasters as unknown as Array<{ themePath?: string }>;
		expect(masters.every((m) => m.themePath === 'ppt/theme/theme1.xml')).toBeTruthy();
		expect(ops.pushHistory).toHaveBeenCalledOnce();
	});

	it('does not apply a theme without a handler or when not editable', async () => {
		const { store, actions } = setup({ handler: null });
		await actions.applyThemeByPath('ppt/theme/theme1.xml', true);
		expect(store.get().dirty).toBeFalsy();

		const { handler, setSlideMasterTheme } = themePartHandler();
		const locked = setup({ handler, editable: false });
		await locked.actions.applyThemeByPath('ppt/theme/theme1.xml', true);
		expect(setSlideMasterTheme).not.toHaveBeenCalled();
		expect(locked.ops.pushHistory).not.toHaveBeenCalled();
	});

	it('re-colours templateElementsBySlideId alongside slides when editing the theme', async () => {
		const OFFICE_ACCENT1 = '#4472C4';
		const ION_ACCENT1 = '#B01513';
		const officeColors: PptxThemeColorScheme = {
			dk1: '#000000',
			lt1: '#FFFFFF',
			dk2: '#44546A',
			lt2: '#E7E6E6',
			accent1: OFFICE_ACCENT1,
			accent2: '#ED7D31',
			accent3: '#A5A5A5',
			accent4: '#FFC000',
			accent5: '#5B9BD5',
			accent6: '#70AD47',
			hlink: '#0563C1',
			folHlink: '#954F72',
		};
		const templateShape = {
			type: 'shape',
			id: 'bg_1',
			x: 0,
			y: 0,
			width: 200,
			height: 100,
			shapeStyle: { fillColor: OFFICE_ACCENT1 },
		} as PptxElement;

		const { store, actions } = setup();
		store.set({
			colorScheme: officeColors,
			templateElementsBySlideId: { s1: [templateShape] },
		});

		await actions.applyThemeEdit({
			colorScheme: { ...officeColors, accent1: ION_ACCENT1 },
			fontScheme: { majorFont: { latin: 'Calibri' }, minorFont: { latin: 'Calibri' } },
			name: 'Custom',
		});

		const patched = store.get().templateElementsBySlideId.s1?.[0] as {
			shapeStyle?: { fillColor?: string };
		};
		expect(patched?.shapeStyle?.fillColor).toBe(ION_ACCENT1);
	});

	it('patches the active slide with history (theme override)', () => {
		const { store, ops, actions } = setup();
		actions.updateActiveSlide({ clrMapOverride: { bg1: 'lt1' } });
		expect(ops.pushHistory).toHaveBeenCalledOnce();
		expect(store.get().slides[0].clrMapOverride).toStrictEqual({ bg1: 'lt1' });
		expect(store.get().dirty).toBeTruthy();
	});

	it('merges presentation-setting patches into the current properties', () => {
		const { store, ops, actions } = setup();
		store.set({ presentationProperties: { loopContinuously: true } });
		actions.updatePresentationSettings({ showType: 'kiosk' });
		expect(ops.updatePresentationProperties).toHaveBeenCalledWith({
			loopContinuously: true,
			showType: 'kiosk',
		});
	});

	/**
	 * The SLIDE BACKGROUND card's template rows: React/Vue/Angular's shortcut
	 * to edit a layout/master's background colour directly from the slide
	 * inspector, without leaving the slide for Master Views.
	 */
	describe('setTemplateBackground / getTemplateBackgroundColor', () => {
		it('writes through the handler and mirrors the colour back onto the store', () => {
			const setTemplateBackground = vi.fn();
			const handler = { setTemplateBackground } as unknown as PptxHandler;
			const { store, actions } = setup({ handler });

			actions.setTemplateBackground('ppt/slideMasters/slideMaster1.xml', '#ff0000');

			expect(setTemplateBackground).toHaveBeenCalledWith(
				'ppt/slideMasters/slideMaster1.xml',
				'#ff0000',
			);
			const masters = store.get().slideMasters as unknown as Array<{
				path: string;
				backgroundColor?: string;
			}>;
			expect(masters[0].backgroundColor).toBe('#ff0000');
			expect(masters[1].backgroundColor).toBeUndefined();
			expect(store.get().dirty).toBeTruthy();
		});

		it('does nothing without a handler or while not editable', () => {
			const setTemplateBackground = vi.fn();
			const handler = { setTemplateBackground } as unknown as PptxHandler;

			const noHandler = setup();
			noHandler.actions.setTemplateBackground('ppt/slideMasters/slideMaster1.xml', '#ff0000');
			expect(setTemplateBackground).not.toHaveBeenCalled();

			const locked = setup({ handler, editable: false });
			locked.actions.setTemplateBackground('ppt/slideMasters/slideMaster1.xml', '#ff0000');
			expect(setTemplateBackground).not.toHaveBeenCalled();
		});

		it('reads the colour straight from the handler', () => {
			const getTemplateBackgroundColor = vi.fn().mockReturnValue('#123456');
			const handler = { getTemplateBackgroundColor } as unknown as PptxHandler;
			const { actions } = setup({ handler });

			expect(actions.getTemplateBackgroundColor('ppt/slideMasters/slideMaster1.xml')).toBe(
				'#123456',
			);
			expect(getTemplateBackgroundColor).toHaveBeenCalledWith('ppt/slideMasters/slideMaster1.xml');
		});

		it('returns undefined without a handler', () => {
			const { actions } = setup();
			expect(
				actions.getTemplateBackgroundColor('ppt/slideMasters/slideMaster1.xml'),
			).toBeUndefined();
		});
	});

	describe('table style DEFINITION editor ("Edit style...")', () => {
		it('updateTableStyleMap replaces the map the renderer reads and marks the deck dirty', () => {
			const { store, actions } = setup();
			store.set({ tableStyleMap: { a: { styleId: 'a', styleName: 'a' } } });

			const nextMap = {
				a: { styleId: 'a', styleName: 'a' },
				b: { styleId: 'b', styleName: 'b' },
			};
			actions.updateTableStyleMap(nextMap);

			expect(store.get().tableStyleMap).toStrictEqual(nextMap);
			expect(store.get().tableStylesToDelete).toStrictEqual([]);
			expect(store.get().dirty).toBeTruthy();
		});

		it('updateTableStyleMap drops a pending delete when the id reappears', () => {
			const { store, actions } = setup();
			store.set({
				tableStyleMap: { a: { styleId: 'a', styleName: 'a' } },
				tableStylesToDelete: ['b'],
			});

			actions.updateTableStyleMap({
				a: { styleId: 'a', styleName: 'a' },
				b: { styleId: 'b', styleName: 'b' },
			});

			expect(store.get().tableStylesToDelete).toStrictEqual([]);
		});

		it('deleteTableStyle removes the entry and records the id for save-time deletion', () => {
			const { store, actions } = setup();
			store.set({
				tableStyleMap: {
					a: { styleId: 'a', styleName: 'a' },
					b: { styleId: 'b', styleName: 'b' },
				},
			});

			actions.deleteTableStyle('a');

			expect(store.get().tableStyleMap).toStrictEqual({ b: { styleId: 'b', styleName: 'b' } });
			expect(store.get().tableStylesToDelete).toStrictEqual(['a']);
			expect(store.get().dirty).toBeTruthy();
		});

		it('ignores table style edits while not editable', () => {
			const { store, actions } = setup({ editable: false });
			const before = store.get().tableStyleMap;

			actions.updateTableStyleMap({ a: { styleId: 'a', styleName: 'a' } });
			actions.deleteTableStyle('a');

			expect(store.get().tableStyleMap).toBe(before);
			expect(store.get().dirty).toBeFalsy();
		});
	});
});
