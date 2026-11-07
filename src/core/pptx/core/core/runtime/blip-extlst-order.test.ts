/**
 * A picture effect added on save goes before the blip's extension list
 * (`CT_Blip`: effects, then `a:extLst`), also on a loaded picture whose blip
 * already has one, and both load back.
 */
import { readFileSync } from 'node:fs';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { ImagePptxElement, PptxImageEffects } from '../../types';

const FIXTURE = new URL('../../../__tests__/fixtures/picture-fixture.pptx', import.meta.url);

async function saveWithEffects(effects: PptxImageEffects | undefined) {
	const handler = new PptxHandler();
	const data = await handler.load(readFileSync(FIXTURE));
	const picture = data.slides[0]?.elements.find(
		(el) => el.type === 'picture' || el.type === 'image',
	) as ImagePptxElement | undefined;
	if (picture && effects) {
		picture.imageEffects = { ...picture.imageEffects, ...effects };
	}
	const bytes = await handler.save(data.slides);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const reloaded = await new PptxHandler().load(bytes);
	return {
		blip: /<a:blip\b[\s\S]*?<\/a:blip>/u.exec(xml)?.[0] ?? '',
		effects: (
			reloaded.slides[0]?.elements.find((el) => el.type === 'picture' || el.type === 'image') as
				| ImagePptxElement
				| undefined
		)?.imageEffects,
	};
}

describe('blip extension list order', () => {
	it('writes new effects before the existing a:extLst and reads them back', async () => {
		const { blip, effects } = await saveWithEffects({
			alphaModFix: 50,
			lum: { bright: 20, contrast: -10 },
			grayscale: true,
		});
		expect(blip).toMatch(/<\/a:extLst><\/a:blip>$/u);
		const extLst = blip.indexOf('<a:extLst>');
		for (const effect of ['<a:alphaModFix', '<a:lum', '<a:grayscl']) {
			expect(blip.indexOf(effect)).toBeGreaterThan(0);
			expect(blip.indexOf(effect)).toBeLessThan(extLst);
		}
		expect(blip).toContain('a14:useLocalDpi');
		expect(effects?.alphaModFix).toBe(50);
		expect(effects?.lum).toEqual({ bright: 20, contrast: -10 });
	});

	it('leaves an unedited picture blip as it was', async () => {
		const { blip } = await saveWithEffects(undefined);
		expect(blip).toMatch(
			/^<a:blip r:embed="rId\d+"><a:extLst><a:ext uri="\{28A0092B-C50C-407E-A947-70E740481C1C\}">/u,
		);
		expect(blip).toMatch(/<\/a:extLst><\/a:blip>$/u);
	});
});
