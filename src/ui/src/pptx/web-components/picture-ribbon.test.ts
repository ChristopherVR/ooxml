// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { PptxImageLikeElement } from 'ooxml-core/pptx';
import { registerPptxWebControls } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

const picture: PptxImageLikeElement = {
	id: 'picture',
	type: 'picture',
	x: 5,
	y: 8,
	width: 90,
	height: 65,
	rotation: 340,
};

function mount(family: HTMLElementTagNameMap['pptx-ui-picture-ribbon']['family'], editable = true) {
	const strip = document.createElement('pptx-ui-picture-ribbon');
	const state = {
		element: picture,
		editable,
		translate: (key: string) => key,
		patch: vi.fn(),
		altText: vi.fn(),
		selectionPane: vi.fn(),
	};
	strip.family = family;
	strip.state = state;
	document.body.append(strip);
	function command(id: string) {
		const button = strip
			.querySelector(`[data-ribbon-control="pictureFormat.${id}"]`)
			?.shadowRoot?.querySelector('button');
		if (!button) throw new Error(`Missing command ${id}`);
		return button;
	}
	return { state, command };
}

describe('picture ribbon', () => {
	it('resets picture formatting in one patch and disables raster tools with a reason', () => {
		const { state, command } = mount('picture-adjust');
		command('adjust.resetPicture').click();
		expect(state.patch).toHaveBeenCalledExactlyOnceWith(
			'picture',
			expect.objectContaining({ imageEffects: undefined, cropLeft: 0 }),
		);
		const remove = command('adjust.removeBackground');
		expect(remove.disabled).toBe(true);
		expect(remove.title).toMatch(/Requires a local PNG or JPEG/u);
	});

	it('offers Alt Text for a read-only selection and reports why Picture Layout is disabled', () => {
		const { state, command } = mount('picture-accessibility', false);
		command('accessibility.altText').click();
		expect(state.altText).toHaveBeenCalledOnce();
		const layout = mount('picture-styles').command('pictureStyles.layout');
		expect(layout.disabled).toBe(true);
		expect(layout.title).toMatch(/from a selection/u);
	});

	it('rotates through the element update and opens the existing Selection Pane', () => {
		const { state, command } = mount('picture-arrange');
		command('arrange.rotateRight').click();
		expect(state.patch).toHaveBeenCalledExactlyOnceWith('picture', { rotation: 70 });
		command('arrange.selectionPane').click();
		expect(state.selectionPane).toHaveBeenCalledOnce();
	});
});
