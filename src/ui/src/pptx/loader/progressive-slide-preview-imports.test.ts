// @vitest-environment happy-dom
import { PptxHandler } from 'ooxml-core/pptx';
import { expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createProgressiveSlidePreview } from './progressive-slide-preview';

vi.mock(import('../dom/load-presentation'), () => {
	throw new Error('Opening previews must work without the DOM presentation loader');
});

it('renders a preview without loading the DOM presentation loader', async () => {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	const root = document.createElement('div');
	const host = document.createElement('div');
	host.dataset.pptxLoadPreviewHost = '';
	root.append(host);
	document.body.append(root);
	const preview = createProgressiveSlidePreview({
		getRoot: () => root,
		translate: createTranslator(),
		isCurrent: () => true,
	});
	try {
		const [slide] = data.slides;
		if (!slide) throw new Error('The test deck needs a first slide');
		await preview.show(handler, {
			slide,
			index: 0,
			total: data.slides.length,
			width: data.width,
			height: data.height,
		});
		expect(host.querySelector('[data-pptx-opening-preview]')).not.toBeNull();
	} finally {
		handler.dispose();
		preview.dispose();
		root.remove();
	}
});
