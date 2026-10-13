import type { RibbonGalleryContext, ToolbarActionId } from 'ooxml-ui/pptx';
import { EMPTY_RIBBON_TRANSITION_DRAFT } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../../i18n';
import { createRibbon } from '../ribbon';
import type { RibbonHandlers, RibbonSelectionState } from '../ribbon-types';

/**
 * Picture Format > Size > Crop in the vanilla ribbon: the contextual pane
 * draws the shared Size strip, and its Crop split button runs the same edit
 * handlers as Home > Arrange.
 */

/** Every method access returns a memoised `vi.fn()`. */
function fakeActions<T extends object>(overrides: Partial<T> = {}): T {
	const cache = new Map<string, ReturnType<typeof vi.fn>>();
	return new Proxy({} as T, {
		get(_target, prop) {
			if (typeof prop !== 'string') {
				return undefined;
			}
			if (prop in overrides) {
				return overrides[prop as keyof T];
			}
			let fn = cache.get(prop);
			if (!fn) {
				fn = vi.fn();
				cache.set(prop, fn);
			}
			return fn;
		},
	});
}

function handlers(): RibbonHandlers {
	return {
		nav: fakeActions(),
		primary: fakeActions(),
		file: fakeActions<RibbonHandlers['file']>({ getRecentPresentationsCount: () => 10 }),
		slideShow: fakeActions<RibbonHandlers['slideShow']>({ showOptions: () => ({}) }),
		insert: fakeActions(),
		edit: fakeActions(),
		findReplace: fakeActions(),
		design: fakeActions(),
		transitions: fakeActions<RibbonHandlers['transitions']>({
			readDraft: () => ({ ...EMPTY_RIBBON_TRANSITION_DRAFT }),
			readTransition: () => undefined,
		}),
		draw: fakeActions(),
	};
}

const PICTURE = {
	id: 'pic',
	type: 'picture',
	x: 0,
	y: 0,
	width: 200,
	height: 100,
} as unknown as PptxElement;

function selection(extra: Partial<RibbonSelectionState> = {}): RibbonSelectionState {
	const galleryContext: RibbonGalleryContext = {
		element: PICTURE,
		themeColorMap: { accent1: '#156082', dk1: '#000000', lt1: '#FFFFFF' },
	};
	return { hasClipboard: false, slideCount: 1, galleryContext, canCrop: true, ...extra };
}

function mount(hiddenActions?: ToolbarActionId[]) {
	const h = handlers();
	const ribbon = createRibbon(document, createTranslator(), h, hiddenActions);
	ribbon.setEditState({ editable: true, canUndo: false, canRedo: false });
	ribbon.setEditable(true);
	ribbon.updateSelection(PICTURE, selection());
	ribbon.el
		.querySelector<HTMLButtonElement>('[data-ribbon-contextual-tab="pictureFormat"]')!
		.click();
	const size = (hook: string) =>
		ribbon.el.querySelector<HTMLButtonElement>(
			`[data-ribbon-group="pictureFormat.size"] [data-pptx-ribbon-control="${hook}"]`,
		)!;
	return { ribbon, edit: h.edit, size };
}

describe('picture Format > Size (vanilla)', () => {
	it('toggles crop mode and reflects it as pressed', () => {
		const { ribbon, edit, size } = mount();
		expect(size('crop').closest('[data-pptx-crop-controls]')).not.toBeNull();
		expect(size('crop').getAttribute('aria-pressed')).toBe('false');
		size('crop').click();
		expect(edit.toggleCropMode).toHaveBeenCalledOnce();
		ribbon.updateSelection(PICTURE, selection({ cropActive: true }));
		expect(size('crop').getAttribute('aria-pressed')).toBe('true');
	});

	it('crops to an aspect ratio and fits from the caret menu', () => {
		const { ribbon, edit, size } = mount();
		size('crop-menu').click();
		ribbon.el.querySelector<HTMLElement>('[data-pptx-crop-aspect="1:1"]')!.click();
		expect(edit.cropToAspect).toHaveBeenCalledWith(1, 1);
		size('crop-menu').click();
		ribbon.el.querySelector<HTMLElement>('[data-pptx-crop-action="fit"]')!.click();
		expect(edit.cropFit).toHaveBeenCalledOnce();
	});

	it('is disabled without a croppable picture and hidden when the host hides crop', () => {
		const { ribbon, size } = mount();
		ribbon.updateSelection(PICTURE, selection({ canCrop: false }));
		expect(size('crop').disabled).toBeTruthy();
		const hidden = mount(['crop']);
		expect(hidden.size('crop').hidden).toBeTruthy();
		expect(hidden.size('crop-menu').hidden).toBeTruthy();
	});
});
