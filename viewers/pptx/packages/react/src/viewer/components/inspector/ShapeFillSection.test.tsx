// @vitest-environment happy-dom
/**
 * The React Fill section draws the shared fill descriptors and the compound
 * controls `resolveShapeFillView` asks for, and every change goes out as the
 * shared fill patch.
 */
import type { ShapeStyle } from 'pptx-viewer-core';
import React, { act } from 'react';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createMarkedRoot } from '../../i18n/marking-translator.test-support';
import { ShapeFillSection } from './ShapeFillSection';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createMarkedRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function render(style: ShapeStyle | undefined, onPatch = vi.fn(), onSetFillColor = vi.fn()) {
	act(() => {
		root.render(
			<div className='grid grid-cols-2'>
				<ShapeFillSection
					style={style}
					isLine={false}
					onUpdateShapeStyle={onPatch}
					onSetFillColor={onSetFillColor}
				/>
			</div>,
		);
	});
	return onPatch;
}

const byTestId = (id: string) => container.querySelector<HTMLInputElement>(`[data-testid="${id}"]`);
const all = (id: string) => container.querySelectorAll<HTMLElement>(`[data-testid="${id}"]`);

function change(el: HTMLInputElement, value: string): void {
	act(() => {
		el.value = value;
		el.dispatchEvent(new Event('change', { bubbles: true }));
	});
}

describe('shapeFillSection', () => {
	it('switches the fill kind through the shared mode patch', () => {
		const onPatch = render({ fillMode: 'solid', fillColor: '#ff0000' });
		expect(byTestId('fill-mode')!.value).toBe('solid');
		expect(byTestId('fill-opacity')).not.toBeNull();
		change(byTestId('fill-mode')!, 'pattern');
		expect(onPatch).toHaveBeenCalledWith({
			fillMode: 'pattern',
			fillPatternPreset: 'pct20',
			fillPatternBackgroundColor: '#ffffff',
		});
	});

	it('draws nothing but the kind for no fill and a slide background fill', () => {
		render({ useBackgroundFill: true });
		expect(byTestId('fill-mode')!.value).toBe('background');
		expect(byTestId('fill-opacity')).toBeNull();
		expect(container.querySelector('input[type="color"]')).toBeNull();
	});

	it('lists the gradient stops and adds one through the shared action', () => {
		const onPatch = render({
			fillMode: 'gradient',
			fillGradientStops: [
				{ color: '#000000', position: 0 },
				{ color: '#ffffff', position: 100 },
			],
		});
		expect(all('fill-gradient-stop')).toHaveLength(2);
		expect(byTestId('fill-gradient-angle')!.value).toBe('90');
		const add = [...container.querySelectorAll('button')].find((b) =>
			b.textContent?.includes('pptx.gradient.addStop'),
		)!;
		act(() => add.click());
		expect(onPatch.mock.calls[0]![0].fillGradientStops).toHaveLength(3);
	});

	it('picks a pattern preset from the swatches', () => {
		const onPatch = render({ fillMode: 'pattern', fillPatternPreset: 'pct20' });
		const swatches = all('fill-pattern-swatch');
		expect(swatches).toHaveLength(56);
		expect([...swatches].filter((s) => s.getAttribute('aria-pressed') === 'true')).toHaveLength(1);
		act(() => (swatches[3] as HTMLButtonElement).click());
		expect(onPatch).toHaveBeenCalledWith({ fillMode: 'pattern', fillPatternPreset: 'pct25' });
	});

	it('shows the picture tiling for a picture fill', () => {
		render({ fillMode: 'image', fillImageMode: 'tile' });
		expect(byTestId('fill-image-mode')!.value).toBe('tile');
	});
});
