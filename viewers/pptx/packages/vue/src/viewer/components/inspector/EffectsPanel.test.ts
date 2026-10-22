// oxlint-disable react-hooks/rules-of-hooks
import { mount } from '@vue/test-utils';
import { SHAPE_QUICK_STYLES } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import EffectsPanel from './EffectsPanel.vue';
import { setControlValue } from './test-control-value';

function shape(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		type: 'shape',
		id: 'sp 1',
		x: 100,
		y: 50,
		width: 200,
		height: 120,
		rotation: 0,
		opacity: 1,
		shapeStyle: {},
		...overrides,
	} as PptxElement;
}

function image(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		type: 'chart',
		id: 'ch 1',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		opacity: 1,
		...overrides,
	} as PptxElement;
}

function lastPatch(wrapper: ReturnType<typeof mount>): Partial<PptxElement> {
	const events = wrapper.emitted('update');
	expect(events).toBeTruthy();
	const ev = events as unknown[][];
	return ev[ev.length - 1][0] as Partial<PptxElement>;
}

describe('effectsPanel', () => {
	it('emits a shallow opacity patch (0-100 slider to 0-1)', async () => {
		const wrapper = mount(EffectsPanel, { props: { element: shape() } });
		const range = wrapper.find('input[type="range"]');
		await setControlValue(range, '50');
		expect(lastPatch(wrapper)).toStrictEqual({ opacity: 0.5 });
	});

	it('shows a muted note for non-shape-like elements and hides every effect control', () => {
		const wrapper = mount(EffectsPanel, { props: { element: image() } });
		expect(wrapper.find('.pptx-vue-effects-note').exists()).toBeTruthy();
		expect(wrapper.findAll('pptx-ui-checkbox')).toHaveLength(0);
	});

	it('forwards a shadow-section patch through to the update event untouched', async () => {
		// EffectsPanel composes EffectsShadowSection + EffectsGlowReflectionSection
		// (each split out to stay under the repo's 300-LOC-per-file budget) and
		// re-emits their `update` events as its own. The shadow toggle button lives
		// on the child; this pins that forwarding wire, not the shadow logic itself
		// (covered in `EffectsShadowSection.test.ts`).
		const wrapper = mount(EffectsPanel, { props: { element: shape() } });
		const toggle = wrapper.find('[data-testid="fx-outer-shadow-toggle"]');
		expect(toggle.exists()).toBeTruthy();
		await setControlValue(toggle, true);
		const patch = lastPatch(wrapper) as { shapeStyle?: { shadowColor?: string } };
		expect(patch.shapeStyle?.shadowColor).toBe('#000000');
	});

	it('forwards a glow/reflection-section patch through to the update event untouched', async () => {
		const wrapper = mount(EffectsPanel, { props: { element: shape() } });
		const toggle = wrapper.find('[data-testid="fx-glow-toggle"]');
		expect(toggle.exists()).toBeTruthy();
		await setControlValue(toggle, true);
		const patch = lastPatch(wrapper) as { shapeStyle?: { glowColor?: string } };
		expect(patch.shapeStyle?.glowColor).toBe('#ffff00');
	});

	it('applies a Quick Styles preset merged onto the current shapeStyle', () => {
		const wrapper = mount(EffectsPanel, {
			// `fillOpacity` is not set by any SHAPE_QUICK_STYLES preset, so its
			// survival in the patch is what proves this is a MERGE, not a replace.
			props: { element: shape({ shapeStyle: { fillOpacity: 0.42, strokeWidth: 7 } }) },
		});
		// The tiles are the shared gallery's (`pptx-ui-ribbon-gallery`), keyed by preset.
		wrapper.get<HTMLButtonElement>('[data-gallery-item="blueFill"]').element.click();
		expect(lastPatch(wrapper)).toStrictEqual({
			shapeStyle: { fillOpacity: 0.42, ...SHAPE_QUICK_STYLES[0]!.style },
		});
	});

	it('keeps the Quick Styles tiles unpickable when the deck is read-only', () => {
		const wrapper = mount(EffectsPanel, { props: { element: shape(), canEdit: false } });
		const tiles = wrapper.findAll<HTMLButtonElement>('[data-gallery-item]');
		expect(tiles).toHaveLength(SHAPE_QUICK_STYLES.length);
		expect(tiles.every((tile) => tile.element.disabled)).toBeTruthy();
		tiles[0]!.element.click();
		expect(wrapper.emitted('update')).toBeUndefined();
	});
});
