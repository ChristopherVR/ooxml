// @vitest-fresh-modules: asserts the layout switch before the lazy built-in layout library loads.
import type {
	PptxElement,
	PptxSmartArtData,
	PptxSmartArtNode,
	SmartArtLayoutType,
} from 'pptx-viewer-core';
import { resetSmartArtEditCounter } from 'pptx-viewer-core';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { computed, ref } from 'vue';

import { smartArtLayoutLabel, useSmartArtEditing } from './useSmartArtEditing';

function node(id: string, text: string, parentId?: string): PptxSmartArtNode {
	return { id, text, parentId };
}

function dataRef(initial: PptxSmartArtData) {
	const r = ref<PptxSmartArtData>(initial);
	return { r, computed: computed<PptxSmartArtData>(() => r.value) };
}

function baseData(overrides: Partial<PptxSmartArtData> = {}): PptxSmartArtData {
	return {
		nodes: [node('n1', 'A'), node('n2', 'B'), node('n3', 'C')],
		resolvedLayoutType: 'list',
		colorScheme: 'colorful1',
		style: 'flat',
		...overrides,
	} as PptxSmartArtData;
}

/** Extract the smartArtData out of an emitted element patch. */
function patchData(patch: Partial<PptxElement>): PptxSmartArtData {
	const data = (patch as { smartArtData?: PptxSmartArtData }).smartArtData;
	if (!data) {
		throw new Error('patch has no smartArtData');
	}
	return data;
}

describe('useSmartArtEditing', () => {
	beforeEach(() => {
		resetSmartArtEditCounter();
	});

	it('commit applies a whole new data object', () => {
		const { computed: data } = dataRef(baseData());
		const apply = vi.fn();
		const api = useSmartArtEditing({ smartArtData: data, apply });
		api.commit(baseData({ nodes: [node('n1', 'Only')] }));
		expect(patchData(apply.mock.calls[0][0]).nodes).toHaveLength(1);
	});

	it('setColorScheme skips when unchanged and applies when changed', () => {
		const { computed: data } = dataRef(baseData());
		const apply = vi.fn();
		const api = useSmartArtEditing({ smartArtData: data, apply });
		api.setColorScheme('colorful1');
		expect(apply).not.toHaveBeenCalled();
		api.setColorScheme('monochromatic2');
		expect(patchData(apply.mock.calls[0][0]).colorScheme).toBe('monochromatic2');
	});

	it('switchLayout preserves nodes and updates the resolved layout', () => {
		const { computed: data } = dataRef(baseData());
		const apply = vi.fn();
		const api = useSmartArtEditing({ smartArtData: data, apply });
		api.switchLayout('hierarchy');
		const next = patchData(apply.mock.calls[0][0]);
		expect(next.resolvedLayoutType).toBe('hierarchy');
		expect(next.nodes).toHaveLength(3);
	});

	it('switchLayout drops the previous layout definition so the new layout is drawn', () => {
		const { computed: data } = dataRef(
			baseData({
				layoutDefinition: { uniqueId: 'urn:test/layout/old' },
				presLayoutVars: { direction: 'norm' },
			} as Partial<PptxSmartArtData>),
		);
		const apply = vi.fn();
		const api = useSmartArtEditing({ smartArtData: data, apply });
		api.switchLayout('cycle');
		const next = patchData(apply.mock.calls[0][0]);
		expect(next.layoutDefinition).toBeUndefined();
		expect(next.presLayoutVars).toBeUndefined();
	});

	it('never mutates the input data object', () => {
		const initial = baseData();
		const { computed: data } = dataRef(initial);
		const api = useSmartArtEditing({ smartArtData: data, apply: vi.fn() });
		api.setStyle('intense');
		api.switchLayout('cycle');
		expect(initial.nodes).toHaveLength(3);
		expect(initial.resolvedLayoutType).toBe('list');
	});

	it('smartArtLayoutLabel resolves the shared category key for every layout', () => {
		// The helper hands the catalogue key to `translate`; asserting on the key
		// keeps this test about the wiring rather than about English wording.
		const echo = (key: string): string => key;
		expect(smartArtLayoutLabel('cycle', echo)).toBe('pptx.smartart.category.cycle');
		expect(smartArtLayoutLabel('target', echo)).toBe('pptx.smartart.category.target');
		expect(smartArtLayoutLabel('bending', echo)).toBe('pptx.smartart.category.bending');
	});

	it('smartArtLayoutLabel falls back to the raw token for an unmapped layout', () => {
		// A deck may carry a layout newer than the catalogue; showing the token
		// beats showing nothing, and proves the value set is never filtered.
		expect(smartArtLayoutLabel('unknown' as SmartArtLayoutType, (key) => key)).toBe('unknown');
	});
});
