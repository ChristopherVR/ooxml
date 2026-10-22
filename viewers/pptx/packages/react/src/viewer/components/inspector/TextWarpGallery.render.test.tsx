// @vitest-environment happy-dom
/**
 * The Text Warp section draws its presets on the shared gallery (`pptx-ui-ribbon-gallery`, a
 * compact panel), marks the current one and turns a pick into the text style change.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { TextWarpGallery } from './TextWarpGallery';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

const gallery = () => container.querySelector('pptx-ui-ribbon-gallery')!;
const tile = (preset: string) =>
	container.querySelector<HTMLButtonElement>(`[data-gallery-item="${preset}"]`)!;
const pressed = () =>
	[...container.querySelectorAll('[data-gallery-item][aria-pressed="true"]')].map((el) =>
		el.getAttribute('data-gallery-item'),
	);

describe('textWarpGallery', () => {
	it('keeps every preset on the shared panel, hidden until the section opens', () => {
		act(() => root.render(<TextWarpGallery ts={undefined} onUpdateTextStyle={() => {}} />));
		expect(gallery().getAttribute('mode')).toBe('panel');
		expect(gallery().hasAttribute('compact')).toBeTruthy();
		expect(gallery().closest('[hidden]')).not.toBeNull();
		expect(container.querySelectorAll('[data-gallery-item]')).toHaveLength(39);
		expect(pressed()).toStrictEqual(['textNoShape']);
		act(() => container.querySelector<HTMLButtonElement>('button')!.click());
		expect(gallery().closest('[hidden]')).toBeNull();
	});

	it('marks the text style preset', () => {
		act(() =>
			root.render(
				<TextWarpGallery ts={{ textWarpPreset: 'textArchUp' }} onUpdateTextStyle={() => {}} />,
			),
		);
		expect(pressed()).toStrictEqual(['textArchUp']);
		expect(tile('textArchUp').getAttribute('aria-label')).toBe('Arch Up');
	});

	it('sets the picked preset, and No Transform clears it', () => {
		const update = vi.fn();
		act(() =>
			root.render(
				<TextWarpGallery ts={{ textWarpPreset: 'textArchUp' }} onUpdateTextStyle={update} />,
			),
		);
		act(() => tile('textWave1').click());
		act(() => tile('textNoShape').click());
		expect(update.mock.calls).toStrictEqual([
			[{ textWarpPreset: 'textWave1' }],
			[{ textWarpPreset: undefined }],
		]);
	});
});
