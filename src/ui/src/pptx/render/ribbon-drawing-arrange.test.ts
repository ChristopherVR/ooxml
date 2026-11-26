import { expect, it, vi } from 'vitest';

import {
	drawingArrangeItems,
	flipArrangeSelection,
	runDrawingArrange,
} from './ribbon-drawing-arrange';
import { drawingHomeControls } from './ribbon-home-state-objects';

it('keeps Selection Pane reachable without a selection and routes the editing rows', () => {
	const align = vi.fn();
	const selectionPane = vi.fn();
	const actions = { align, selectionPane, group: vi.fn() };
	const arrange = {
		actions,
		canEdit: true,
		selectedCount: 0,
		selectedElement: null,
		canGroup: false,
	};
	const menu = drawingHomeControls({ editable: true, hasSelection: false, arrange })[
		'home.drawing.arrange'
	];
	expect(menu?.disabled).toBe(false);
	expect([...new Set(menu?.items?.map((entry) => entry.group))]).toEqual([
		'Order Objects',
		'Group Objects',
		'Position Objects',
	]);
	expect(menu?.items?.find((entry) => entry.value === 'group')?.disabled).toBe(true);
	expect(menu?.items?.find((entry) => entry.value === 'left')?.disabled).toBe(true);
	expect(runDrawingArrange('selectionPane', actions)).toBe(true);
	expect(selectionPane).toHaveBeenCalledOnce();
	expect(runDrawingArrange('centerH', actions)).toBe(true);
	expect(align).toHaveBeenCalledExactlyOnceWith('centerH');
	expect(runDrawingArrange('front', actions)).toBe(false);
});

it('reports the selection, edit mode and lock that disable each command', () => {
	const selectedElement = {
		id: 'shape',
		type: 'shape',
		x: 0,
		y: 0,
		width: 30,
		height: 40,
		locks: { noRotation: true },
	} as const;
	const input = {
		actions: {
			group: vi.fn(),
			ungroup: vi.fn(),
			distribute: vi.fn(),
			rotation: () => ({ element: selectedElement, patch: vi.fn() }),
		},
		selectedElement,
		canEdit: true,
		selectedCount: 1,
		canGroup: false,
	};
	const reason = (value: string, changes: Partial<typeof input> = {}) =>
		drawingArrangeItems({ ...input, ...changes }).find((row) => row.value === value)
			?.descriptionKey;

	expect(reason('group')).toBe('pptx.arrange.needsGroupSelection');
	expect(reason('group', { selectedCount: 2 })).toBe('pptx.arrange.objectLocked');
	expect(reason('ungroup')).toBe('pptx.arrange.needsUngroupSelection');
	expect(reason('distribute-horizontal')).toBe('pptx.arrange.needsDistributeSelection');
	expect(reason('rotate-right')).toBe('pptx.arrange.objectLocked');
	expect(reason('front', { canEdit: false })).toBe('pptx.draw.readOnly');
	expect(reason('distribute-horizontal', { selectedCount: 3 })).toBeUndefined();
});

it('flips each selected object on the requested axis and skips missing selections', () => {
	const shape = { id: 'a', type: 'shape', x: 0, y: 0, width: 30, height: 40 } as const;
	const elements = [shape, { ...shape, id: 'b', flipHorizontal: true }, { ...shape, id: 'c' }];
	const patch = vi.fn();
	flipArrangeSelection(elements, ['a', 'missing', 'b'], 'horizontal', patch);
	expect(patch.mock.calls).toStrictEqual([
		['a', { flipHorizontal: true }],
		['b', { flipHorizontal: false }],
	]);
	patch.mockClear();
	flipArrangeSelection(elements, ['b'], 'vertical', patch);
	expect(patch).toHaveBeenCalledExactlyOnceWith('b', { flipVertical: true });
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
