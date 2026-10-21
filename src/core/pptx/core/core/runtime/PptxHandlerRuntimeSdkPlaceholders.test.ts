import { describe, expect, it } from 'vitest';
import { createTextElement } from '../../builders/sdk/ElementFactory';
import { PresentationBuilder } from '../../builders/sdk/PresentationBuilder';
import { PptxHandlerRuntime } from '../PptxHandlerRuntime';

class InspectableRuntime extends PptxHandlerRuntime {
	contextState() {
		return {
			relationships: [...this.slideRelsMap].map(([path, rels]) => [path, [...rels]]),
			parts: Object.keys(this.zip.files),
		};
	}
}

describe('SDK placeholder runtime context', () => {
	it('resolves a layout directly without changing archive parts or relationship maps', async () => {
		const { handler, data } = await PresentationBuilder.create();
		const bytes = await handler.save(data.slides);
		const runtime = new InspectableRuntime();
		await runtime.load(bytes.buffer as ArrayBuffer);
		const before = runtime.contextState();
		const layout = runtime.getLayoutOptions().find((item) => item.name === 'Title Slide')!;
		const element = runtime.createPlaceholderElement(
			layout.path,
			createTextElement('Title'),
			'ctrTitle',
			{},
		);
		expect(element.width).toBeGreaterThan(0);
		expect(element.inheritedTransform).toBeDefined();
		expect(runtime.contextState()).toStrictEqual(before);
		expect(() =>
			runtime.createPlaceholderElement(layout.path, createTextElement('Bad'), 'ctrTitle', {
				x: NaN,
			}),
		).toThrow(/finite/);
		expect(runtime.contextState()).toStrictEqual(before);
		runtime.dispose();
	});
});
