// @vitest-environment happy-dom
/**
 * The React generic field renderer draws the shared shape Effects descriptors:
 * one toggle per section, its fields only while on, and every change goes out
 * as the descriptor's `Partial<ShapeStyle>` patch.
 */
import { SHAPE_EFFECT_SECTIONS, SHAPE_LINE_SECTIONS } from 'ooxml-ui/pptx';
import type { ShapeStyle } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { InspectorFields } from './InspectorFields';

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

function render(style: ShapeStyle | undefined, onPatch: (p: Partial<ShapeStyle>) => void) {
	act(() => {
		root.render(
			<div className='grid grid-cols-2'>
				<InspectorFields sections={SHAPE_EFFECT_SECTIONS} target={style} onPatch={onPatch} />
			</div>,
		);
	});
}

const byTestId = <E extends Element = HTMLInputElement>(id: string) =>
	container.querySelector<E & HTMLInputElement>(`[data-testid="${id}"]`);

describe('inspectorFields (effects)', () => {
	it('draws a toggle per effect section and no fields while every effect is off', () => {
		render({}, vi.fn());
		for (const id of [
			'fx-outer-shadow-toggle',
			'fx-inner-shadow-toggle',
			'fx-glow-toggle',
			'fx-soft-edge-toggle',
			'fx-reflection-toggle',
			'fx-blur-toggle',
			'fx-bevel-toggle',
		]) {
			expect(byTestId(id)?.tagName.toLowerCase()).toBe('pptx-ui-checkbox');
		}
		expect(byTestId('fx-outer-shadow-blur')).toBeNull();
	});

	it('emits the enable patch when a toggle is ticked', () => {
		const onPatch = vi.fn();
		render({}, onPatch);
		const toggle = byTestId('fx-glow-toggle')!;
		act(() => {
			toggle.checked = true;
			toggle.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({
			glowColor: '#ffff00',
			glowRadius: 6,
			glowOpacity: 0.75,
		});
	});

	it("shows an enabled effect's fields and writes a clamped number", () => {
		const onPatch = vi.fn();
		render({ shadowColor: '#112233', shadowBlur: 3, shadowAngle: 0, shadowDistance: 4 }, onPatch);
		const blur = byTestId('fx-outer-shadow-blur')!;
		expect(blur.value).toBe('3');
		expect(blur.getAttribute('aria-label')).toBeTruthy();
		act(() => {
			const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
			setter.call(blur, '500');
			blur.dispatchEvent(new Event('input', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith(expect.objectContaining({ shadowBlur: 96 }));
		expect(byTestId('fx-outer-shadow-rotate-with-shape')).not.toBeNull();
	});
});

describe('inspectorFields (line options)', () => {
	function renderLine(style: ShapeStyle | undefined, onPatch: (p: Partial<ShapeStyle>) => void) {
		act(() => {
			root.render(
				<div className='grid grid-cols-2'>
					<InspectorFields sections={SHAPE_LINE_SECTIONS} target={style} onPatch={onPatch} />
				</div>,
			);
		});
	}

	it('draws the untitled line fields with their current values', () => {
		renderLine({ strokeWidth: 3, strokeOpacity: 0.4, lineCap: 'sq' }, vi.fn());
		expect(byTestId('line-width')!.value).toBe('3');
		expect(byTestId('line-opacity')!.value).toBe('40');
		expect(byTestId('line-cap')!.tagName.toLowerCase()).toBe('pptx-ui-select');
		expect(container.querySelector('pptx-ui-checkbox')).toBeNull();
	});

	it('writes a dash choice as a shape-style patch', () => {
		const onPatch = vi.fn();
		renderLine({}, onPatch);
		const dash = byTestId('line-dash')!;
		act(() => {
			dash.value = 'dash';
			dash.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({ strokeDash: 'dash' });
	});
});
