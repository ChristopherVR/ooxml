import type { PptxElement } from 'ooxml-core/pptx';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { FocusSelectionInput } from './focus-targets';
import {
	AI_TOOL_FLASH_MS,
	createPptxAiPanelController,
	isAiFocusPinned,
	resolveAiEffectiveTargets,
} from './panel-controller';

function setup(selection: FocusSelectionInput, element?: PptxElement) {
	const onOpen = vi.fn();
	const live = { ...selection };
	const controller = createPptxAiPanelController({
		getSelection: () => live,
		getSelectedElement: () => element,
		onOpen,
	});
	return { controller, live, onOpen };
}

describe('createPptxAiPanelController', () => {
	beforeEach(() => {
		vi.useFakeTimers();
	});
	afterEach(() => {
		vi.useRealTimers();
	});

	it('follows the live selection until pinned, and picks win over a pin', () => {
		const { controller, live } = setup({
			activeSlideIndex: 0,
			selectedElementIds: ['a'],
			selectedElementId: 'a',
		});
		expect(controller.getEffectiveTargets()).toEqual([
			{ kind: 'element', slideIndex: 0, elementId: 'a' },
		]);
		controller.pinFocus();
		live.selectedElementIds = [];
		live.selectedElementId = null;
		expect(isAiFocusPinned(controller.getSnapshot())).toBe(true);
		expect(controller.getEffectiveTargets()[0]).toMatchObject({ elementId: 'a' });

		controller.addPick(2, 'z');
		controller.addPick(2, 'z');
		expect(controller.getSnapshot().pickTargets).toHaveLength(1);
		expect(controller.getEffectiveTargets()).toEqual([
			{ kind: 'element', slideIndex: 2, elementId: 'z' },
		]);
		expect(isAiFocusPinned(controller.getSnapshot())).toBe(false);
		controller.clearPicks();
		controller.clearPinnedFocus();
		expect(controller.getEffectiveTargets()).toEqual([{ kind: 'slide', slideIndex: 0 }]);
	});

	it('opens with a fresh prefill nonce for ask and a fix directive for fix', () => {
		const element = { id: 'e1', type: 'shape' } as PptxElement;
		const { controller, onOpen } = setup(
			{ activeSlideIndex: 1, selectedElementIds: ['e1'], selectedElementId: 'e1' },
			element,
		);
		controller.askAboutSelection();
		expect(controller.getSnapshot()).toMatchObject({
			isOpen: true,
			prefill: { text: '', nonce: 1 },
		});
		controller.fixSelection();
		expect(controller.getSnapshot().prefill).toEqual({
			text: 'Review and fix any issues with this shape (id=e1) on slide 2.',
			nonce: 2,
		});
		expect(onOpen).toHaveBeenCalledTimes(2);
		controller.toggle();
		expect(controller.getSnapshot().isOpen).toBe(false);
	});

	it('draws pick rings and a transient tool ring that clears after the flash window', () => {
		const { controller } = setup({
			activeSlideIndex: 0,
			selectedElementIds: [],
			selectedElementId: null,
		});
		controller.startPicking();
		expect(controller.getSnapshot().pickMode).toBe(true);
		controller.addPick(0, 'p');
		controller.flashToolTarget({ slideIndex: 0, elementIds: ['t'] });
		expect(controller.getSnapshot().canvasHighlights).toEqual([
			{ slideIndex: 0, elementId: 'p', variant: 'pick' },
			{ slideIndex: 0, elementId: 't', variant: 'active' },
		]);
		expect(controller.getSnapshot().canvasAnimating).toBe(true);
		vi.advanceTimersByTime(AI_TOOL_FLASH_MS);
		expect(controller.getSnapshot().canvasHighlights).toHaveLength(1);
		controller.clearPicks();
		expect(controller.getSnapshot()).toMatchObject({ canvasAnimating: false, pickMode: false });
	});

	it('animates for a slide-level tool call with no element ring', () => {
		const { controller } = setup({
			activeSlideIndex: 0,
			selectedElementIds: [],
			selectedElementId: null,
		});
		controller.flashToolTarget(null);
		expect(controller.getSnapshot()).toMatchObject({ canvasAnimating: true, canvasHighlights: [] });
		controller.dispose();
	});
});

describe('resolveAiEffectiveTargets', () => {
	it('falls back to the live targets', () => {
		const live = [{ kind: 'slide' as const, slideIndex: 3 }];
		expect(resolveAiEffectiveTargets({ pinnedFocus: null, pickTargets: [] }, live)).toBe(live);
	});
});
