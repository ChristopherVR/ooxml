import { expect, it, vi } from 'vitest';

import { runDrawingArrange } from './ribbon-drawing-arrange';
import { drawingHomeControls } from './ribbon-home-state-objects';

it('keeps Selection Pane reachable without a selection and routes the editing rows', () => {
	const align = vi.fn();
	const selectionPane = vi.fn();
	const actions = { align, selectionPane, group: vi.fn() };
	const arrange = {
		actions,
		canEdit: true,
		selected: false,
		canGroup: false,
		canUngroup: false,
		canDistribute: false,
	};
	const menu = drawingHomeControls({ editable: true, hasSelection: false, arrange })[
		'home.drawing.arrange'
	];
	expect(menu?.disabled).toBe(false);
	expect(menu?.items?.find((entry) => entry.value === 'group')?.disabled).toBe(true);
	expect(menu?.items?.find((entry) => entry.value === 'left')?.disabled).toBe(true);
	expect(runDrawingArrange('selectionPane', actions)).toBe(true);
	expect(selectionPane).toHaveBeenCalledOnce();
	expect(runDrawingArrange('centerH', actions)).toBe(true);
	expect(align).toHaveBeenCalledExactlyOnceWith('centerH');
	expect(runDrawingArrange('front', actions)).toBe(false);
});

it('rotates the selected object in either direction and honors its rotation lock', () => {
	const patch = vi.fn();
	const element = {
		id: 'shape',
		type: 'shape',
		x: 0,
		y: 0,
		width: 30,
		height: 40,
		rotation: 45,
	} as const;
	const actions = { rotation: () => ({ element, patch }) };
	runDrawingArrange('rotate-right', actions);
	runDrawingArrange('rotate-left', actions);
	expect(patch.mock.calls).toStrictEqual([
		['shape', { rotation: 135 }],
		['shape', { rotation: 315 }],
	]);
	patch.mockClear();
	runDrawingArrange('rotate-right', {
		rotation: () => ({ element: { ...element, locks: { noRotation: true } }, patch }),
	});
	expect(patch).not.toHaveBeenCalled();
});
