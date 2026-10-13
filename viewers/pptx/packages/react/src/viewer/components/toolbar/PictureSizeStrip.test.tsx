import type { ToolbarActionId } from 'ooxml-ui/pptx';
// @vitest-environment happy-dom
/**
 * Picture Format > Size > Crop in the React ribbon: the contextual tab draws
 * the shared Size strip, and its Crop split button drives the same crop
 * controller as Home > Arrange.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { PictureCropController } from '../../hooks/usePictureCropMode';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { ContextualTabSection } = await import('./ContextualTabSection');
const { ShapeFormatContext } = await import('../shape-format-context');

function cropStub(overrides: Partial<PictureCropController> = {}): PictureCropController {
	return {
		element: null,
		canCrop: true,
		toggle: vi.fn(),
		enter: vi.fn(),
		commit: vi.fn(),
		cancel: vi.fn(),
		liveUpdate: vi.fn(),
		cropToAspect: vi.fn(),
		fill: vi.fn(),
		fit: vi.fn(),
		...overrides,
	};
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function render(crop: PictureCropController, hiddenActions?: ToolbarActionId[]): void {
	const value = { canMergeShapes: false, mergeShapes: () => {}, crop };
	act(() =>
		root.render(
			<ShapeFormatContext.Provider value={value}>
				<ContextualTabSection tab='pictureFormat' hiddenActions={hiddenActions} />
			</ShapeFormatContext.Provider>,
		),
	);
}

const sizeControl = (hook: string) =>
	container.querySelector<HTMLButtonElement>(
		`[data-ribbon-group="pictureFormat.size"] [data-pptx-ribbon-control="${hook}"]`,
	);

describe('picture Format > Size', () => {
	it('toggles crop mode from its Crop button and shows it pressed', () => {
		const crop = cropStub();
		render(crop);
		expect(sizeControl('crop')?.getAttribute('aria-pressed')).toBe('false');
		act(() => sizeControl('crop')?.click());
		expect(crop.toggle).toHaveBeenCalledOnce();

		render(cropStub({ element: { id: 'pic' } as never }));
		expect(sizeControl('crop')?.getAttribute('aria-pressed')).toBe('true');
	});

	it('crops to an aspect ratio and fits from the caret menu', () => {
		const crop = cropStub();
		render(crop);
		act(() => sizeControl('crop-menu')?.click());
		act(() => document.querySelector<HTMLElement>('[data-pptx-crop-aspect="1:1"]')?.click());
		expect(crop.cropToAspect).toHaveBeenCalledWith(1, 1);
		act(() => sizeControl('crop-menu')?.click());
		act(() => document.querySelector<HTMLElement>('[data-pptx-crop-action="fit"]')?.click());
		expect(crop.fit).toHaveBeenCalledOnce();
	});

	it('is disabled without a croppable picture and hidden when the host hides crop', () => {
		render(cropStub({ canCrop: false }));
		expect(sizeControl('crop')?.disabled).toBeTruthy();
		render(cropStub(), ['crop']);
		expect(sizeControl('crop')?.hidden).toBeTruthy();
		expect(sizeControl('crop-menu')?.hidden).toBeTruthy();
	});
});
