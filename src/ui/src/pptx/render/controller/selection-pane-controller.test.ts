import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import {
	createSelectionPaneController,
	formatSelectionPaneLabel,
	selectionPaneLabel,
	selectionPaneRows,
} from './selection-pane-controller';

const el = (partial: Partial<PptxElement>): PptxElement =>
	({ id: 'e', type: 'shape', x: 0, y: 0, width: 1, height: 1, ...partial }) as PptxElement;

const t = (key: string, params?: Record<string, string | number>): string =>
	key === 'pptx.selectionPane.elementLabel'
		? `${params?.type} #${params?.number}`
		: (({ 'pptx.elementType.smartArt': 'Diagram' } as Record<string, string>)[key] ?? key);

function setup(canEdit = true) {
	const deps = { canEdit: () => canEdit, onRename: vi.fn(), onReorder: vi.fn() };
	return { controller: createSelectionPaneController(deps), deps };
}

describe('selection pane labels', () => {
	it('names a row by its name, then its text, then its translated type and position', () => {
		expect(selectionPaneLabel(el({ name: '  Logo ' }), 0)).toStrictEqual({
			kind: 'name',
			text: 'Logo',
		});
		const long = 'x'.repeat(40);
		expect(selectionPaneLabel(el({ type: 'text', text: ` ${long}` } as never), 0)).toStrictEqual({
			kind: 'text',
			text: long.slice(0, 32),
		});
		expect(formatSelectionPaneLabel(selectionPaneLabel(el({ type: 'smartArt' }), 2), t)).toBe(
			'Diagram #3',
		);
		expect(selectionPaneLabel(el({ type: 'weird' as never }), 0)).toMatchObject({
			typeKey: 'pptx.elementType.object',
		});
	});

	it('lists the rows top-most first with their stacking position', () => {
		const rows = selectionPaneRows([el({ id: 'back' }), el({ id: 'front' })]);
		expect(rows.map((row) => [row.element.id, row.index])).toStrictEqual([
			['front', 1],
			['back', 0],
		]);
	});
});

describe('selection pane controller', () => {
	it('writes an edited rename and nothing for an unedited one', () => {
		const { controller, deps } = setup();
		controller.startRename('e1', 'Shape 3');
		expect(controller.commitRename()).toBe(false);
		controller.startRename('e1', 'Shape 3');
		controller.setRenameDraft('Logo');
		expect(controller.commitRename()).toBe(true);
		expect(deps.onRename).toHaveBeenCalledExactlyOnceWith('e1', 'Logo');
		// The commit's own blur re-enters: a no-op.
		expect(controller.commitRename()).toBe(false);
	});

	it('clears a name when the box is emptied, and cancels on Escape', () => {
		const { controller, deps } = setup();
		controller.startRename('e1', 'Logo');
		controller.setRenameDraft('');
		controller.commitRename();
		expect(deps.onRename).toHaveBeenCalledWith('e1', '');
		controller.startRename('e1', 'Logo');
		controller.cancelRename();
		expect(controller.getSnapshot().renamingId).toBeNull();
	});

	it('reorders by dragging and clears the drag state', () => {
		const { controller, deps } = setup();
		controller.dragStart(2);
		controller.dragOver(2);
		expect(controller.getSnapshot().dragOverIndex).toBeNull();
		controller.dragOver(0);
		expect(controller.getSnapshot()).toMatchObject({ dragIndex: 2, dragOverIndex: 0 });
		controller.drop(0);
		expect(deps.onReorder).toHaveBeenCalledExactlyOnceWith(2, 0);
		expect(controller.getSnapshot()).toMatchObject({ dragIndex: null, dragOverIndex: null });
		controller.dragStart(1);
		controller.drop(1);
		expect(deps.onReorder).toHaveBeenCalledOnce();
	});

	it('refuses renaming and dragging on a read-only deck', () => {
		const { controller, deps } = setup(false);
		expect(controller.startRename('e1', 'Shape 1')).toBe(false);
		expect(controller.dragStart(0)).toBe(false);
		controller.drop(1);
		expect(deps.onRename).not.toHaveBeenCalled();
		expect(deps.onReorder).not.toHaveBeenCalled();
	});
});
