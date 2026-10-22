import { afterEach, describe, expect, it, vi } from 'vitest';

import { destroyMounted, mount } from './region-customization.test-support';
import { allButtons } from './ui/dialog-footer.test-support';

afterEach(destroyMounted);

/**
 * `onPropertiesChange` (issue 151, item 5), through a mounted viewer: the document-properties
 * dialog reports what the user saved, with the same payload as the other bindings.
 */
describe('onPropertiesChange', () => {
	function saveTitle(options: { editable: boolean; onPropertiesChange: () => void }): void {
		const { viewer } = mount(options);
		(viewer as unknown as { openDocumentProperties(): void }).openDocumentProperties();
		const dialog = document.querySelector<HTMLElement>('.pptxv-props-overlay [role="dialog"]');
		expect(dialog).not.toBeNull();
		const title = dialog!.querySelector<HTMLInputElement>('input');
		title!.value = 'Quarterly review';
		title!.dispatchEvent(new Event('input', { bubbles: true }));
		allButtons(dialog!)
			.find((button) => button.textContent === 'Save')
			?.click();
	}

	it('reports the saved core properties', () => {
		const onPropertiesChange = vi.fn();

		saveTitle({ editable: true, onPropertiesChange });

		expect(onPropertiesChange).toHaveBeenCalledExactlyOnceWith(
			expect.objectContaining({ title: 'Quarterly review' }),
		);
	});

	it('stays silent for a read-only viewer', () => {
		const onPropertiesChange = vi.fn();

		saveTitle({ editable: false, onPropertiesChange });

		expect(onPropertiesChange).not.toHaveBeenCalled();
	});
});
