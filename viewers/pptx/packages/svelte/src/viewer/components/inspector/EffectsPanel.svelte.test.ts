import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import EffectsPanel from './EffectsPanel.svelte';

/**
 * EffectsPanel tests: enable/disable + field edits for outer shadow, inner
 * shadow, glow, reflection, and soft edge, drawn by the generic
 * `InspectorFields` renderer from the shared `SHAPE_EFFECT_SECTIONS`. Named `*.svelte.test.ts` so `mountPanel`'s props object can be
 * wrapped in `$state(...)` and refreshed between chained interactions.
 */

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function shapeEl(shapeStyle: Record<string, unknown> = {}): PptxElement {
	return {
		type: 'shape',
		id: 's1',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		shapeStyle,
	} as unknown as PptxElement;
}

function makeEditor(el: PptxElement): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [el] }]);
	editor.select(el.id);
	return editor;
}

function currentEl(editor: EditorState): PptxElement {
	const el = editor.slides[0]?.elements[0];
	if (!el) {
		throw new Error('element missing');
	}
	return el;
}

function mountPanel(editor: EditorState, el: PptxElement) {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const props = $state({ editor, el });
	const instance = mount(EffectsPanel, { target, props });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return {
		target,
		setProps: (next: { el: PptxElement }) => {
			Object.assign(props, next);
			flushSync();
		},
	};
}

function toggle(target: HTMLElement, testId: string): HTMLInputElement | null {
	return target.querySelector<HTMLInputElement>(`[data-testid="${testId}"]`);
}

describe('effectsPanel', () => {
	it('enables outer shadow with defaults and reveals its fields', () => {
		const editor = makeEditor(shapeEl());
		const { target, setProps } = mountPanel(editor, currentEl(editor));
		expect(target.querySelector('.pptx-svelte-fields-grid')).toBeNull();

		toggle(target, 'fx-outer-shadow-toggle')?.click();
		flushSync();
		setProps({ el: currentEl(editor) });

		const el = currentEl(editor) as {
			shapeStyle?: { shadowColor?: string; shadowOpacity?: number };
		};
		expect(el.shapeStyle?.shadowColor).toBe('#000000');
		expect(el.shapeStyle?.shadowOpacity).toBe(0.35);
		expect(target.querySelector('.pptx-svelte-fields-grid')).not.toBeNull();
	});

	it('disables outer shadow back to transparent', () => {
		const editor = makeEditor(shapeEl({ shadowColor: '#ff0000', shadowOpacity: 0.5 }));
		const { target, setProps } = mountPanel(editor, currentEl(editor));

		toggle(target, 'fx-outer-shadow-toggle')?.click();
		flushSync();
		setProps({ el: currentEl(editor) });

		const el = currentEl(editor) as { shapeStyle?: { shadowColor?: string } };
		expect(el.shapeStyle?.shadowColor).toBe('transparent');
	});

	it('enables glow and updates its radius', () => {
		const editor = makeEditor(shapeEl());
		const { target, setProps } = mountPanel(editor, currentEl(editor));

		toggle(target, 'fx-glow-toggle')?.click();
		flushSync();
		setProps({ el: currentEl(editor) });

		let el = currentEl(editor) as { shapeStyle?: { glowColor?: string; glowRadius?: number } };
		expect(el.shapeStyle?.glowColor).toBe('#ffff00');
		expect(el.shapeStyle?.glowRadius).toBe(6);

		const radiusInput = target.querySelector<HTMLInputElement>('[data-testid="fx-glow-radius"]');
		if (!radiusInput) {
			throw new Error('glow radius input not found');
		}
		radiusInput.value = '12';
		radiusInput.dispatchEvent(new Event('change', { bubbles: true }));
		flushSync();
		setProps({ el: currentEl(editor) });

		el = currentEl(editor) as { shapeStyle?: { glowRadius?: number } };
		expect(el.shapeStyle?.glowRadius).toBe(12);
	});

	it('enables reflection with defaults', () => {
		const editor = makeEditor(shapeEl());
		const { target, setProps } = mountPanel(editor, currentEl(editor));

		toggle(target, 'fx-reflection-toggle')?.click();
		flushSync();
		setProps({ el: currentEl(editor) });

		const el = currentEl(editor) as {
			shapeStyle?: { reflectionBlurRadius?: number; reflectionStartOpacity?: number };
		};
		expect(el.shapeStyle?.reflectionBlurRadius).toBe(3);
		expect(el.shapeStyle?.reflectionStartOpacity).toBe(50);
	});

	it('enables soft edge with a non-zero default radius', () => {
		const editor = makeEditor(shapeEl());
		const { target, setProps } = mountPanel(editor, currentEl(editor));

		toggle(target, 'fx-soft-edge-toggle')?.click();
		flushSync();
		setProps({ el: currentEl(editor) });

		const el = currentEl(editor) as { shapeStyle?: { softEdgeRadius?: number } };
		expect(el.shapeStyle?.softEdgeRadius).toBeGreaterThan(0);
	});

	it('draws the blur and top bevel sections from the shared descriptors', () => {
		const editor = makeEditor(shapeEl({ shape3d: { bevelTopType: 'circle' } }));
		const { target } = mountPanel(editor, currentEl(editor));

		expect(toggle(target, 'fx-blur-toggle')).not.toBeNull();
		const bevelType = target.querySelector<HTMLSelectElement>('[data-testid="fx-bevel-type"]');
		expect(bevelType?.localName).toBe('pptx-ui-select');
		expect(target.querySelector('[data-testid="fx-bevel-width"]')).not.toBeNull();
	});
});
