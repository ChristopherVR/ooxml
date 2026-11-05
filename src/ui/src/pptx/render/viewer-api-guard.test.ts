import { describe, expect, it, vi } from 'vitest';

import {
	GATED_VIEWER_EDIT_METHODS,
	gateViewerEdits,
	resolveReportedViewerMode,
} from './viewer-api-guard';

function api() {
	return {
		addSlide: vi.fn(),
		deleteSlides: vi.fn(),
		duplicateSlides: vi.fn(),
		moveSlide: vi.fn(),
		toggleHideSlides: vi.fn(),
		updateElement: vi.fn(),
		deleteElements: vi.fn(),
		duplicateElement: vi.fn(() => 'copy'),
		getSlideCount: vi.fn(() => 3),
	};
}

describe('gateViewerEdits', () => {
	it('calls through while editing is allowed, with the original arguments', () => {
		const base = api();
		const gated = gateViewerEdits(base, () => true);
		gated.addSlide(1);
		gated.updateElement('a', { x: 1 } as never);
		expect(base.addSlide).toHaveBeenCalledWith(1);
		expect(base.updateElement).toHaveBeenCalledWith('a', { x: 1 });
		expect(gated.duplicateElement('a')).toBe('copy');
	});

	it('still lets a read-only viewer without any slide create its first one', () => {
		const base = { ...api(), getSlideCount: vi.fn(() => 0) };
		const gated = gateViewerEdits(base, () => false);
		gated.addSlide();
		gated.deleteSlides([0]);
		expect(base.addSlide).toHaveBeenCalledOnce();
		expect(base.deleteSlides).not.toHaveBeenCalled();
	});

	it('turns every edit into a no-op while editing is refused, and reads the right on each call', () => {
		const base = api();
		let allowed = false;
		const gated = gateViewerEdits(base, () => allowed);
		for (const name of GATED_VIEWER_EDIT_METHODS) {
			expect((gated[name] as (...args: unknown[]) => unknown)(0, 0)).toBeUndefined();
			expect(base[name]).not.toHaveBeenCalled();
		}
		// Members that are not edits stay as they were.
		expect(gated.getSlideCount()).toBe(3);
		allowed = true;
		gated.deleteSlides([0]);
		expect(base.deleteSlides).toHaveBeenCalledOnce();
	});
});

describe('resolveReportedViewerMode', () => {
	it.each([
		['edit', true, 'edit'],
		['edit', false, 'preview'],
		['preview', true, 'preview'],
		['present', false, 'present'],
		['master', false, 'master'],
	] as const)('%s with canEdit %s reports %s', (mode, canEdit, expected) => {
		expect(resolveReportedViewerMode(mode, canEdit)).toBe(expected);
	});
});
