import JSZip from 'jszip';

import { renameFixtureParts } from '../../../pptx/__tests__/helpers/rename-package-parts';
import { PresentationBuilder } from '../../../pptx/core/builders/sdk';
import { PptxHandler } from '../../../pptx/core/PptxHandler';
import type { PptxElement } from '../../../pptx/core/types';
import type { SecurityCase } from '../case';

/** What an id may hold once it can reach CSS selectors, attributes and markup. */
const SAFE_ID = /^[A-Za-z0-9\-._~/%]+$/u;
const MARKUP = '"></style><i data-probe="1">';

function ids(elements: readonly PptxElement[] | undefined, into: string[] = []): string[] {
	for (const element of elements ?? []) {
		into.push(element.id);
		if (element.type === 'group') ids(element.children, into);
	}
	return into;
}

export const cases: readonly SecurityCase[] = [
	{
		name: 'markup in a slide part name never reaches the ids of the elements on that slide',
		limits: { ms: 5_000, heapMb: 256 },
		async run() {
			const { handler, data, createSlide } = await PresentationBuilder.create();
			data.slides = [createSlide('Blank').addText('Kept text').build()];
			const saved = await handler.save(data.slides);
			handler.dispose();
			const zip = await JSZip.loadAsync(saved);
			const slide = Object.keys(zip.files).find((path) => /^ppt\/slides\/[^/]+\.xml$/u.test(path))!;
			const bytes = await renameFixtureParts(
				saved,
				new Map([[slide, `ppt/slides/s${MARKUP}.xml`]]),
			);
			const reader = new PptxHandler();
			const loaded = await reader.load(bytes.slice().buffer);
			reader.dispose();
			const found = ids(loaded.slides[0]?.elements);
			return { elements: found.length > 0, unsafeIds: found.filter((id) => !SAFE_ID.test(id)) };
		},
		expected: { elements: true, unsafeIds: [] },
	},
];
