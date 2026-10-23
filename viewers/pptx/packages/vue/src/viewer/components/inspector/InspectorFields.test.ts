// oxlint-disable react-hooks/rules-of-hooks
import { mount } from '@vue/test-utils';
import { SHAPE_EFFECT_SECTIONS } from 'ooxml-ui/pptx';
import type { ShapeStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import InspectorFields from './InspectorFields.vue';
import { setControlValue } from './test-control-value';

/**
 * The Vue generic field renderer draws the shared shape Effects descriptors:
 * one toggle per section, its fields only while on, and every change goes out
 * as the descriptor's `Partial<ShapeStyle>` patch.
 */
function mountFields(target: ShapeStyle | undefined) {
	return mount(InspectorFields, { props: { sections: SHAPE_EFFECT_SECTIONS, target } });
}

function lastPatch(wrapper: ReturnType<typeof mountFields>): Partial<ShapeStyle> {
	const events = wrapper.emitted('patch') as unknown[][] | undefined;
	expect(events).toBeTruthy();
	return events![events!.length - 1]![0] as Partial<ShapeStyle>;
}

describe('inspectorFields (effects)', () => {
	it('draws a toggle per effect section and no fields while every effect is off', () => {
		const wrapper = mountFields({});
		for (const id of [
			'fx-outer-shadow-toggle',
			'fx-inner-shadow-toggle',
			'fx-glow-toggle',
			'fx-soft-edge-toggle',
			'fx-reflection-toggle',
			'fx-blur-toggle',
			'fx-bevel-toggle',
		]) {
			expect(wrapper.find(`[data-testid="${id}"]`).exists()).toBeTruthy();
		}
		expect(wrapper.find('input').exists()).toBeFalsy();
	});

	it('emits the enable patch when a toggle is ticked', async () => {
		const wrapper = mountFields({});
		await setControlValue(wrapper.find('[data-testid="fx-soft-edge-toggle"]'), true);
		expect(lastPatch(wrapper)).toStrictEqual({ softEdgeRadius: 6 });
	});

	it("shows an enabled effect's fields and writes a clamped number", async () => {
		const wrapper = mountFields({ innerShadowColor: '#123456', innerShadowBlur: 2 });
		const blur = wrapper.find('[data-testid="fx-inner-shadow-blur"]');
		expect((blur.element as HTMLInputElement).value).toBe('2');
		await setControlValue(blur, '-4');
		expect(lastPatch(wrapper)).toMatchObject({ innerShadowBlur: 0 });
		await setControlValue(wrapper.find('[data-testid="fx-inner-shadow-opacity"]'), '20');
		expect(lastPatch(wrapper)).toMatchObject({ innerShadowOpacity: 0.2 });
	});

	it('writes a bevel type from its drop-down', async () => {
		const wrapper = mountFields({ shape3d: { bevelTopType: 'circle' } });
		await setControlValue(wrapper.find('[data-testid="fx-bevel-type"]'), 'slope');
		expect(lastPatch(wrapper).shape3d).toMatchObject({ bevelTopType: 'slope' });
	});
});
