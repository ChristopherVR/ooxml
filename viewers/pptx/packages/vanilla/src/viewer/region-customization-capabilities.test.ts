// @vitest-own-worker: mounts several full viewers; see region-customization.test-support.ts.
import { afterEach, describe, expect, it } from 'vitest';

import { destroyMounted, mount } from './region-customization.test-support';

afterEach(destroyMounted);

describe('vanilla capabilities.edit', () => {
	it('makes the viewer read-only live and restores it', () => {
		const { container, viewer } = mount({ editable: true });
		const root = (): Element | null => container.querySelector('.pptxv');
		expect(root()?.classList.contains('pptxv-editable')).toBe(true);
		viewer.setCapability('edit', false);
		expect(root()?.classList.contains('pptxv-editable')).toBe(false);
		expect(viewer.getMode()).toBe('preview');
		viewer.setCapability('edit', undefined);
		expect(root()?.classList.contains('pptxv-editable')).toBe(true);
		expect(viewer.getMode()).toBe('edit');
	});

	it('starts read-only when the option sets the capability off', () => {
		const { viewer } = mount({ editable: true, customization: { capabilities: { edit: false } } });
		expect(viewer.getMode()).toBe('preview');
	});

	it('ANDs the legacy readOnly flag', () => {
		const { viewer } = mount({ editable: true, readOnly: true });
		expect(viewer.getMode()).toBe('preview');
	});
});
