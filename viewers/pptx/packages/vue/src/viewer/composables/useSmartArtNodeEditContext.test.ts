// oxlint-disable react-hooks/rules-of-hooks, react/function-component-definition -- Vue setup, not React
import { mount } from '@vue/test-utils';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import { defineComponent, h } from 'vue';

import { injectSmartArtNodeEdit } from './smartart-node-edit';
import type { SmartArtNodeEditContext } from './smartart-node-edit';
import { useSmartArtNodeEditContext } from './useSmartArtNodeEditContext';

const LINKED = 'https://linked-content.test/node.png';

/** A SmartArt element whose drawn node shape has a picture fill linked to the web. */
const model = {
	id: 'art',
	type: 'smartArt',
	x: 0,
	y: 0,
	width: 400,
	height: 200,
	smartArtData: {
		layout: 'basicBlockList',
		style: 'flat',
		nodes: [{ id: 'n1', text: 'One' }],
		drawingShapes: [
			{ id: 's1', modelId: 'n1', x: 0, y: 0, width: 100, height: 50, fillImageUrl: LINKED },
		],
	},
} as unknown as PptxElement;

function provideContext(updateElement: (id: string, patch: Partial<PptxElement>) => void) {
	let context: SmartArtNodeEditContext | undefined;
	const Child = defineComponent({
		setup() {
			context = injectSmartArtNodeEdit();
			return () => null;
		},
	});
	const Root = defineComponent({
		setup() {
			useSmartArtNodeEditContext({
				canEdit: () => true,
				canEditInline: () => true,
				findActiveElement: (id) => (id === model.id ? model : undefined),
				updateElement,
			});
			return () => h(Child);
		},
	});
	const wrapper = mount(Root);
	return { wrapper, context: context! };
}

describe('useSmartArtNodeEditContext', () => {
	it('fills a node from the model element, so a drawn copy without linked content never reaches it', () => {
		const updateElement = vi.fn();
		const { wrapper, context } = provideContext(updateElement);
		context.commitNodeFill?.('art', 'n1', '#00ff00');
		expect(updateElement).toHaveBeenCalledOnce();
		const [id, patch] = updateElement.mock.calls[0]!;
		expect(id).toBe('art');
		expect(JSON.stringify(patch)).toContain(LINKED);
		expect(patch.smartArtData.nodes[0].style.fillColor).toBe('#00ff00');
		wrapper.unmount();
	});

	it('ignores a node fill for an element that is not SmartArt', () => {
		const updateElement = vi.fn();
		const { wrapper, context } = provideContext(updateElement);
		context.commitNodeFill?.('missing', 'n1', '#00ff00');
		expect(updateElement).not.toHaveBeenCalled();
		wrapper.unmount();
	});
});
