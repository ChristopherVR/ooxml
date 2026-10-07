// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { officeGradientPresetFill } from 'ooxml-core/diagram';
import { createTestContext } from './commands/test-support';
import { createGradientTypeField } from './chart-series-gradient-type';

afterEach(() => document.body.replaceChildren());

it('retains imported types and guards disabled or unsupported type picks', () => {
	const ctx = createTestContext();
	let picked = '';
	const field = createGradientTypeField(ctx, (type) => {
		picked = type;
	});
	const input = field.element.querySelector('select')!;
	const fill = officeGradientPresetFill(1);
	fill.path = 'circle';
	field.refresh(fill, false);
	expect(input.value).toBe('circle');
	expect(input.selectedOptions[0]!.textContent).toBe('Radial');
	expect(fill.path).toBe('circle');
	input.dispatchEvent(new Event('change'));
	expect(picked).toBe('');
	fill.path = 'shape';
	field.refresh(fill, false);
	expect(input.value).toBe('shape');
	fill.path = 'future';
	field.refresh(fill, false);
	expect(input.value).toBe('unsupported');
	input.value = 'rect';
	input.dispatchEvent(new Event('change'));
	expect(picked).toBe('rect');
	field.refresh(fill, true);
	input.value = 'linear';
	input.dispatchEvent(new Event('change'));
	expect(picked).toBe('rect');
	ctx.t = (key) => (key === 'Linear gradient' ? '线性' : key);
	field.refresh(fill, false);
	expect(input.options[0]!.textContent).toBe('线性');
});
