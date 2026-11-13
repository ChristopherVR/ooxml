import { describe, expect, it, vi } from 'vitest';

import { describeBuildingBlockActions } from '../building-block-actions';
import { describeBuildingBlockToolbar, runBuildingBlockToolbarControl } from './toolbar';

const availability = describeBuildingBlockActions({
	ready: true,
	hasActiveSlide: true,
	canEdit: true,
	editing: true,
	canInsert: true,
	hasSelection: false,
	selectionHasHyperlink: false,
	selectionHasEquation: false,
	layoutCount: 1,
	hasCurrentLayout: true,
	hasHandler: true,
});

describe('describeBuildingBlockToolbar', () => {
	it('lists the groups in order and disables what cannot run', () => {
		const groups = describeBuildingBlockToolbar({
			availability,
			canUndo: false,
			canRedo: true,
			findOpen: true,
			slideIndex: 0,
			slideCount: 2,
		});
		expect(groups.map((group) => group.id)).toEqual([
			'history',
			'insert',
			'editing',
			'view',
			'navigation',
		]);
		const all = groups.flatMap((group) => group.controls);
		const byId = Object.fromEntries(all.map((control) => [control.id, control]));
		expect(byId.undo?.disabled).toBe(true);
		expect(byId.redo?.disabled).toBe(false);
		// No selection: the link control needs one.
		expect(byId.hyperlink?.disabled).toBe(true);
		expect(byId.insertImage?.disabled).toBe(false);
		expect(byId.find).toEqual({
			id: 'find',
			labelKey: 'pptx.toolbar.findAndReplace',
			disabled: false,
			pressed: true,
		});
		expect(byId.previousSlide?.disabled).toBe(true);
		expect(byId.nextSlide?.disabled).toBe(false);
	});

	it('drops controls the host hides', () => {
		const groups = describeBuildingBlockToolbar({
			availability,
			canUndo: true,
			canRedo: true,
			findOpen: false,
			slideIndex: 0,
			slideCount: 1,
			hiddenActions: ['insert', 'zoom', 'undo'],
		});
		expect(groups.map((group) => group.id)).toEqual(['history', 'editing', 'navigation']);
		expect(groups[0]?.controls.map((control) => control.id)).toEqual(['redo']);
	});
});

describe('runBuildingBlockToolbarControl', () => {
	it('raises requests for inserts and runs view operations', () => {
		const session = {
			request: vi.fn(() => true),
			resetSlide: vi.fn(async () => ({ ok: true as const })),
		};
		const view = {
			undo: vi.fn(),
			redo: vi.fn(),
			toggleFind: vi.fn(),
			zoomOut: vi.fn(),
			zoomIn: vi.fn(),
			zoomToFit: vi.fn(),
			previousSlide: vi.fn(),
			nextSlide: vi.fn(),
		};
		runBuildingBlockToolbarControl('insertEquation', session, view);
		runBuildingBlockToolbarControl('insertImage', session, view);
		runBuildingBlockToolbarControl('resetLayout', session, view);
		runBuildingBlockToolbarControl('find', session, view);
		runBuildingBlockToolbarControl('nextSlide', session, view);
		expect(session.request.mock.calls).toEqual([
			[{ kind: 'equation', existingOmml: null }],
			[{ kind: 'image' }],
		]);
		expect(session.resetSlide).toHaveBeenCalledOnce();
		expect(view.toggleFind).toHaveBeenCalledOnce();
		expect(view.nextSlide).toHaveBeenCalledOnce();
	});
});
