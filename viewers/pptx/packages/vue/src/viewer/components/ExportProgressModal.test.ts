import { mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { describe, expect, it } from 'vitest';

import ExportProgressModal from './ExportProgressModal.vue';

/** The modal renders inside the shared element's open shadow root. */
function inner(wrapper: VueWrapper, selector: string): HTMLElement | null {
	return (wrapper.element as HTMLElement).shadowRoot?.querySelector<HTMLElement>(selector) ?? null;
}

describe('exportProgressModal', () => {
	it('is the shared progress dialog, in place rather than teleported', () => {
		const wrapper = mount(ExportProgressModal, {
			props: { open: true, title: 'Export as PDF', progress: 33, statusMessage: 'Slide 1 of 3' },
			attachTo: document.body,
		});
		expect((wrapper.element as HTMLElement).localName).toBe('pptx-ui-export-progress');
		expect(inner(wrapper, '[role="dialog"]')?.getAttribute('aria-modal')).toBe('true');
		expect(inner(wrapper, '#title')?.textContent).toBe('Export as PDF');
		expect(inner(wrapper, '[role="progressbar"]')?.getAttribute('aria-valuenow')).toBe('33');
		expect(inner(wrapper, '#status')?.textContent).toBe('Slide 1 of 3');
		wrapper.unmount();
	});

	it('emits cancel from Cancel and draws nothing while closed', async () => {
		const wrapper = mount(ExportProgressModal, {
			props: { open: true, title: 'Export', progress: 0 },
		});
		expect(inner(wrapper, '#status')?.textContent).toBe('Processing…');
		inner(wrapper, '[data-testid="pptx-export-progress-cancel"]')!.click();
		expect(wrapper.emitted('cancel')).toHaveLength(1);
		await wrapper.setProps({ open: false });
		expect(inner(wrapper, '[role="dialog"]')).toBeNull();
	});
});
