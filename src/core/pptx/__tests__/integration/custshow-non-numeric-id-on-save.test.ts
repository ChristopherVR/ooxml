import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';

/**
 * `p:custShow/@id` is an `xsd:unsignedInt`. Editors that mint UUID ids used to write them out
 * unchanged, which PowerPoint repairs or rejects. The save now gives such shows the next free
 * number and keeps `p:showPr` and `ppaction://customshow?id=` references pointing at them.
 */
describe('custom show ids on save', () => {
	it('writes valid unsigned-int ids and keeps every reference consistent', async () => {
		const created = await PresentationBuilder.create();
		created.data.slides.push(created.createSlide('Blank').addText('One').build());
		created.data.slides.push(created.createSlide('Blank').addText('Two').build());
		const base = await created.handler.save(created.data.slides);
		const baseXml = await (
			await JSZip.loadAsync(base)
		)
			.file('ppt/presentation.xml')!
			.async('string');
		const rIds = [...baseXml.matchAll(/<p:sldId\s+id="[^"]+"\s+r:id="([^"]+)"/g)].map((m) => m[1]!);
		expect(rIds).toHaveLength(2);

		const slides = created.data.slides.map((slide, i) =>
			i === 0
				? {
						...slide,
						elements: slide.elements.map((el, j) =>
							j === 0
								? { ...el, actionClick: { action: 'ppaction://customshow?id=show-b&return=true' } }
								: el,
						),
					}
				: slide,
		);
		const bytes = await created.handler.save(slides, {
			customShows: [
				{ id: '7', name: 'Seven', slideRIds: [rIds[0]!] },
				{ id: 'show-b', name: 'Uuid', slideRIds: [rIds[1]!] },
				{ id: '7', name: 'Dup', slideRIds: [rIds[0]!] },
			],
			presentationProperties: { showSlidesMode: 'customShow', showSlidesCustomShowId: 'show-b' },
		});

		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file('ppt/presentation.xml')!.async('string');
		const ids = [...xml.matchAll(/<p:custShow\s[^>]*?id="([^"]*)"/g)].map((m) => m[1]);
		const lstIds = [...xml.matchAll(/<p:custShow\b[^>]*\bid="([^"]*)"[^>]*>\s*<p:sldLst/g)].map(
			(m) => m[1],
		);
		expect(lstIds).toEqual(['7', '8', '9']);
		expect(ids.every((id) => /^\d+$/.test(id!))).toBe(true);

		const props = await zip.file('ppt/presProps.xml')!.async('string');
		expect(props).toMatch(/<p:custShow\s+id="8"/);

		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(slideXml).toContain('customshow?id=8');
		expect(slideXml).not.toContain('show-b');
	});
});
