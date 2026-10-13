/**
 * `p:presentation/@firstSlideNum` sets the number a `slidenum` field shows on
 * the first slide (PowerPoint numbers slide `i` as `firstSlideNum + i`). The
 * loader has to keep it and `FieldContextService` has to pass it on, or a deck
 * starting at 45 shows 1 on its first slide.
 */
import { Injector } from '@angular/core';
import JSZip from 'jszip';
import { substituteFieldText } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { FieldContextService } from './field-context.service';
import { LoadContentService } from './load-content.service';

async function deckStartingAt(first: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	try {
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const xml = await zip.file('ppt/presentation.xml')!.async('string');
		zip.file(
			'ppt/presentation.xml',
			xml.replace('<p:presentation ', `<p:presentation firstSlideNum="${first}" `),
		);
		return await zip.generateAsync({ type: 'uint8array' });
	} finally {
		handler.dispose();
	}
}

describe('fieldContextService start slide number', () => {
	it('numbers slide 1 as 45 and slide 2 as 46 for firstSlideNum="45"', async () => {
		const injector = Injector.create({
			providers: [
				{ provide: LoadContentService, useClass: LoadContentService },
				{ provide: FieldContextService, useClass: FieldContextService },
			],
		});
		try {
			const loader = injector.get(LoadContentService);
			await loader.load(await deckStartingAt(45));
			expect(loader.firstSlideNumber()).toBe(45);
			const fields = injector.get(FieldContextService);
			const [first, second] = loader.slides();
			expect(substituteFieldText('1', 'slidenum', fields.forSlide(first))).toBe('45');
			expect(substituteFieldText('2', 'slidenum', fields.forSlide(second))).toBe('46');
		} finally {
			injector.destroy();
		}
	}, 30_000);
});
