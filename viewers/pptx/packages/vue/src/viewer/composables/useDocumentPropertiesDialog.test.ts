import type { PptxAppProperties, PptxCoreProperties, PptxCustomProperty } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';

import { useDocumentPropertiesDialog } from './useDocumentPropertiesDialog';

describe('useDocumentPropertiesDialog', () => {
	it('reports the saved core properties to the host (properties-change)', () => {
		const onCorePropertiesChanged = vi.fn();
		const coreProperties = ref<PptxCoreProperties | undefined>({ title: 'Old' });
		const dialog = useDocumentPropertiesDialog({
			coreProperties,
			customProperties: ref<PptxCustomProperty[]>([]),
			appProperties: ref<PptxAppProperties | undefined>(undefined),
			onCorePropertiesChanged,
		});
		dialog.propertiesOpen.value = true;

		dialog.onPropertiesSave({ core: { title: 'New' }, custom: [] });

		expect(coreProperties.value).toStrictEqual({ title: 'New' });
		expect(onCorePropertiesChanged).toHaveBeenCalledExactlyOnceWith({ title: 'New' });
		expect(dialog.propertiesOpen.value).toBeFalsy();
	});
});
