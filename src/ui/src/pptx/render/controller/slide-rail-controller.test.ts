import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createSlideRailController, slideRailMenuEntries } from './slide-rail-controller';

const ids = ['a', 'b', 'c', 'd'];
const slides = ids.map(
	(id, index) =>
		({ id, rId: id, slideNumber: index + 1, elements: [], hidden: id === 'c' }) as PptxSlide,
);
const plain = { ctrlKey: false, metaKey: false, shiftKey: false };

describe('slide rail controller', () => {
	it('selects, toggles with Ctrl and extends with Shift from the anchor', () => {
		const rail = createSlideRailController({ canEdit: () => true });
		rail.click(plain, 'b', ids);
		rail.click({ ...plain, shiftKey: true }, 'd', ids);
		expect(rail.getSnapshot().selectedIds).toStrictEqual(['b', 'c', 'd']);
		rail.click({ ...plain, ctrlKey: true }, 'c', ids);
		expect(rail.getSnapshot().selectedIds).toStrictEqual(['b', 'd']);
	});

	it('acts on the whole selection when right-clicked inside it, else on that slide', () => {
		const rail = createSlideRailController({ canEdit: () => true });
		rail.setSelection(['b', 'c']);
		rail.openMenu(5, 6, 2, ids);
		expect(rail.getSnapshot().menu).toStrictEqual({
			x: 5,
			y: 6,
			index: 2,
			selectedIndexes: [1, 2],
		});
		rail.openMenu(5, 6, 0, ids);
		expect(rail.getSnapshot().menu?.selectedIndexes).toStrictEqual([0]);
	});

	it('refuses the menu on a read-only deck', () => {
		const rail = createSlideRailController({ canEdit: () => false });
		expect(rail.openMenu(0, 0, 0, ids)).toBe(false);
		expect(rail.getSnapshot().menu).toBeNull();
	});

	it('runs a command on the menu slides and closes the menu', () => {
		const rail = createSlideRailController({ canEdit: () => true });
		rail.setSelection(['b', 'c']);
		rail.openMenu(5, 6, 1, ids);
		const commands = {
			duplicateSlides: vi.fn(),
			openLayoutForSlide: vi.fn(),
			addSlideAfter: vi.fn(),
		};
		rail.run('duplicate', commands);
		expect(commands.duplicateSlides).toHaveBeenCalledWith([1, 2]);
		expect(rail.getSnapshot().menu).toBeNull();
		rail.run('new-slide', commands);
		expect(commands.addSlideAfter).not.toHaveBeenCalled();
		rail.openMenu(5, 6, 1, ids);
		rail.run('layout', commands);
		expect(commands.openLayoutForSlide).toHaveBeenCalledWith(1, 5, 6);
	});

	it('offers Show for an all-hidden selection and refuses deleting every slide', () => {
		const rail = createSlideRailController({ canEdit: () => true });
		rail.openMenu(0, 0, 2, ids);
		const entries = slideRailMenuEntries(rail.getSnapshot(), slides);
		expect(entries.map((entry) => entry.id)).toContain('hide');
		rail.setSelection(ids);
		rail.openMenu(0, 0, 0, ids);
		const remove = slideRailMenuEntries(rail.getSnapshot(), slides).find(
			(entry) => entry.id === 'delete',
		);
		expect(remove?.disabled).toBe(true);
	});
});
