import { describe, expect, it, vi } from 'vitest';

import type { FormatSectionView } from '../inspector/format-sections';
import { createFormatPaneController, formatSectionOpen } from './format-pane-controller';

const fill: FormatSectionView = { id: 'fill', labelKey: 'f', defaultOpen: true };
const effects: FormatSectionView = { id: 'effects', labelKey: 'e', defaultOpen: false };
const size: FormatSectionView = { id: 'sizePosition', labelKey: 's', defaultOpen: false };

describe('createFormatPaneController', () => {
	it('keeps open state per category and section, defaulting to the view', () => {
		const pane = createFormatPaneController();
		expect(pane.isOpen('shape', fill)).toBe(true);
		expect(pane.isOpen('shape', effects)).toBe(false);
		const listener = vi.fn();
		const unsubscribe = pane.subscribe(listener);
		pane.setOpen('shape', 'fill', false);
		expect(pane.isOpen('shape', fill)).toBe(false);
		expect(pane.isOpen('text', fill)).toBe(true);
		expect(formatSectionOpen(pane.getSnapshot(), 'shape', fill)).toBe(false);
		const snapshot = pane.getSnapshot();
		pane.setOpen('shape', 'fill', false);
		expect(pane.getSnapshot()).toBe(snapshot);
		expect(listener).toHaveBeenCalledTimes(1);
		unsubscribe();
	});

	it('keeps state per controller instance', () => {
		const a = createFormatPaneController();
		const b = createFormatPaneController();
		a.setOpen('shape', 'fill', false);
		expect(b.isOpen('shape', fill)).toBe(true);
	});

	it('shows the pane on request and opens the requested section when taken', () => {
		const showPane = vi.fn();
		const pane = createFormatPaneController({ showPane });
		expect(pane.takeRequest('shape', [fill, size])).toBeNull();
		pane.requestFormatPane('sizePosition', { focus: true });
		expect(showPane).toHaveBeenCalledTimes(1);
		expect(pane.getSnapshot().request).toMatchObject({ section: 'sizePosition', focus: true });
		expect(pane.takeRequest('shape', [fill, size])).toEqual({
			section: 'sizePosition',
			focus: true,
		});
		expect(pane.isOpen('shape', size)).toBe(true);
		expect(pane.getSnapshot().request).toBeNull();
		expect(pane.takeRequest('shape', [fill, size])).toBeNull();
	});

	it('falls back to the first section when the requested one is not shown', () => {
		const pane = createFormatPaneController();
		pane.requestFormatPane('effects');
		expect(pane.takeRequest('table', [size])).toEqual({ section: 'sizePosition', focus: false });
		pane.requestFormatPane();
		expect(pane.takeRequest('slide', [])).toBeNull();
	});

	it('reports category changes after the first', () => {
		const pane = createFormatPaneController();
		expect(pane.noteCategory('shape')).toBe(false);
		expect(pane.noteCategory('shape')).toBe(false);
		expect(pane.noteCategory('text')).toBe(true);
	});

	it('takes a late showPane dependency', () => {
		const pane = createFormatPaneController();
		const showPane = vi.fn();
		pane.setShowPane(showPane);
		pane.requestFormatPane('fill');
		expect(showPane).toHaveBeenCalledTimes(1);
	});
});
