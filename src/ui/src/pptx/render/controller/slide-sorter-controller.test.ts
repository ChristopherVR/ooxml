import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createSlideSorterController, slideSorterView } from './slide-sorter-controller';
import type { SlideSorterControllerDeps, SlideSorterKeyEvent } from './slide-sorter-controller';

const slide = (id: string, hidden = false): PptxSlide =>
	({ id, rId: id, elements: [], hidden }) as unknown as PptxSlide;

function setup(overrides: Partial<SlideSorterControllerDeps> = {}) {
	let slides = [slide('a'), slide('b'), slide('c', true), slide('d')];
	let canEdit = true;
	const deps = {
		slides: () => slides,
		activeIndex: () => 1,
		canEdit: () => canEdit,
		onOperation: vi.fn(),
		onPaste: vi.fn(),
		onMove: vi.fn(),
		onClose: vi.fn(),
		...overrides,
	};
	const controller = createSlideSorterController(deps);
	return {
		controller,
		deps,
		setSlides: (next: PptxSlide[]) => (slides = next),
		setCanEdit: (value: boolean) => (canEdit = value),
	};
}

function key(init: Partial<SlideSorterKeyEvent> & { key: string }): SlideSorterKeyEvent {
	return {
		target: null,
		preventDefault: vi.fn(),
		stopPropagation: vi.fn(),
		...init,
	};
}

describe('slide sorter controller', () => {
	it('starts on the active slide and extends and toggles the selection', () => {
		const { controller } = setup();
		expect(controller.getSnapshot().sorter.selectedIds).toEqual(['b']);
		controller.select(3, { shiftKey: true });
		expect(controller.getSnapshot().sorter.selectedIds).toEqual(['b', 'c', 'd']);
		controller.select(2, { ctrlKey: true });
		expect(controller.selectedIndexes()).toEqual([1, 3]);
	});

	it('keeps the snapshot identity when nothing changes', () => {
		const { controller } = setup();
		const before = controller.getSnapshot();
		controller.closeMenu();
		controller.dragLeave();
		controller.setZoom(100);
		expect(controller.getSnapshot()).toBe(before);
	});

	it('opens the menu over a selection it belongs to, and never on a read-only deck', () => {
		const { controller, setCanEdit } = setup();
		controller.select(0, {});
		controller.select(2, { shiftKey: true });
		expect(controller.openMenu(1, 10, 20)).toBe(true);
		expect(controller.getSnapshot().menu).toEqual({ x: 10, y: 20, index: 1 });
		expect(controller.getSnapshot().sorter.selectedIds).toEqual(['a', 'b', 'c']);
		controller.closeMenu();
		setCanEdit(false);
		expect(controller.openMenu(3, 0, 0)).toBe(false);
		expect(controller.getSnapshot().menu).toBeNull();
	});

	it('reports deck operations in descending order and closes the menu', () => {
		const { controller, deps } = setup();
		controller.select(0, {});
		controller.select(2, { shiftKey: true });
		controller.openMenu(0, 0, 0);
		controller.run('duplicate');
		expect(deps.onOperation).toHaveBeenLastCalledWith('duplicate', [2, 1, 0]);
		expect(controller.getSnapshot().menu).toBeNull();
		controller.run('toggle-hidden');
		expect(deps.onOperation).toHaveBeenLastCalledWith('toggle-hidden', [2, 1, 0]);
		controller.run('delete');
		expect(deps.onOperation).toHaveBeenLastCalledWith('delete', [2, 1, 0]);
		expect(controller.getSnapshot().sorter.selectedIds).toEqual([]);
	});

	it('refuses to delete every slide', () => {
		const { controller, deps } = setup();
		controller.run('selectAll');
		controller.run('delete');
		expect(deps.onOperation).not.toHaveBeenCalled();
	});

	it('copies and pastes after the selection', () => {
		const { controller, deps } = setup();
		controller.run('copy');
		expect(slideSorterView(controller.getSnapshot(), deps.slides()).menuContext.hasClipboard).toBe(
			true,
		);
		controller.run('paste');
		expect(deps.onPaste).toHaveBeenCalledTimes(1);
		expect(deps.onPaste.mock.calls[0][0].insertAt).toBe(2);
		expect(controller.pasteInsertAt()).toBe(2);
	});

	it('maps keys through the shared keymap: Escape collapses, then closes', () => {
		const { controller, deps } = setup();
		controller.select(3, { shiftKey: true });
		const escape = key({ key: 'Escape' });
		expect(controller.handleKeydown(escape)).toBe(true);
		expect(escape.preventDefault).toHaveBeenCalled();
		expect(escape.stopPropagation).toHaveBeenCalled();
		expect(controller.getSnapshot().sorter.selectedIds).toEqual(['b']);
		expect(deps.onClose).not.toHaveBeenCalled();
		controller.handleKeydown(key({ key: 'Escape' }));
		expect(deps.onClose).toHaveBeenCalledTimes(1);
	});

	it('leaves unmapped keys and deck-writing keys on a read-only deck alone', () => {
		const { controller, deps, setCanEdit } = setup();
		const plain = key({ key: 'x' });
		expect(controller.handleKeydown(plain)).toBe(false);
		expect(plain.preventDefault).not.toHaveBeenCalled();
		setCanEdit(false);
		expect(controller.handleKeydown(key({ key: 'Delete' }))).toBe(false);
		expect(controller.handleKeydown(key({ key: 'd', ctrlKey: true }))).toBe(false);
		expect(deps.onOperation).not.toHaveBeenCalled();
	});

	it('zooms with the keys and the slider within the sorter range', () => {
		const { controller, deps } = setup();
		controller.handleKeydown(key({ key: '+', ctrlKey: true }));
		expect(controller.getSnapshot().sorter.zoom).toBe(110);
		controller.zoomBy(-1);
		controller.setZoom(1000);
		expect(controller.getSnapshot().sorter.zoom).toBe(200);
		expect(slideSorterView(controller.getSnapshot(), deps.slides()).gridColumns).toBe(2);
	});

	it('routes undo and redo only when the host passes them', () => {
		const onUndo = vi.fn();
		const onRedo = vi.fn();
		const { controller } = setup({ onUndo, onRedo });
		controller.handleKeydown(key({ key: 'z', ctrlKey: true }));
		controller.handleKeydown(key({ key: 'Z', ctrlKey: true, shiftKey: true }));
		controller.handleKeydown(key({ key: 'y', metaKey: true }));
		expect(onUndo).toHaveBeenCalledTimes(1);
		expect(onRedo).toHaveBeenCalledTimes(2);
		expect(setup().controller.handleKeydown(key({ key: 'z', ctrlKey: true }))).toBe(false);
	});

	it('closes the menu on any key', () => {
		const { controller } = setup();
		controller.openMenu(0, 0, 0);
		controller.handleKeydown(key({ key: 'x' }));
		expect(controller.getSnapshot().menu).toBeNull();
	});

	it('drags a tile onto another and clears the drag state', () => {
		const { controller, deps } = setup();
		controller.dragStart(0);
		controller.dragOver(2);
		expect(controller.getSnapshot()).toMatchObject({ dragIndex: 0, dragOverIndex: 2 });
		controller.drop(2);
		expect(deps.onMove).toHaveBeenCalledWith(0, 2);
		expect(controller.getSnapshot()).toMatchObject({ dragIndex: null, dragOverIndex: null });
		controller.drop(1, 3);
		expect(deps.onMove).toHaveBeenLastCalledWith(3, 1);
		controller.drop(1, 1);
		controller.drop(1, Number.NaN);
		expect(deps.onMove).toHaveBeenCalledTimes(2);
	});

	it('never moves a slide on a read-only deck', () => {
		const { controller, deps, setCanEdit } = setup();
		setCanEdit(false);
		controller.dragStart(0);
		controller.drop(2);
		controller.move(1, 0);
		expect(deps.onMove).not.toHaveBeenCalled();
		setCanEdit(true);
		controller.move(0, 9);
		expect(deps.onMove).not.toHaveBeenCalled();
	});

	it('resumes a saved sorter state', () => {
		const first = setup().controller;
		first.run('copy');
		first.setZoom(150);
		const resumed = createSlideSorterController(setup().deps, {
			initial: first.getSnapshot().sorter,
		});
		expect(resumed.getSnapshot().sorter.zoom).toBe(150);
		expect(resumed.getSnapshot().sorter.clipboardIds).toEqual(['b']);
	});

	it('resets onto the active slide once the deck is readable', () => {
		const { controller } = setup();
		controller.reset({ selectedIds: [], anchor: 0, clipboardIds: [], zoom: 100 });
		controller.openMenu(0, 0, 0);
		controller.reset();
		expect(controller.getSnapshot()).toMatchObject({
			sorter: { selectedIds: ['b'], anchor: 1 },
			menu: null,
		});
	});

	it('derives the menu entries and selected ids for the view', () => {
		const { controller, deps } = setup();
		controller.select(2, {});
		const view = slideSorterView(controller.getSnapshot(), deps.slides());
		expect(view.selectedIds.has('c')).toBe(true);
		expect(view.menuEntries.map((entry) => entry.id)).toEqual([
			'copy',
			'duplicate',
			'toggle-hidden',
			'delete',
		]);
		expect(view.zoomScale).toBe(1);
	});
});
