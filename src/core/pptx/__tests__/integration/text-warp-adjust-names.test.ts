/**
 * The two-guide WordArt presets (the waves and the pours) name their first
 * adjustment `adj1`, every other preset `adj`. A wave's `adj1` used to be
 * dropped on load (only `adj` was read), so an edited slide lost the author's
 * amplitude and the renderer painted the default one.
 */
import { describe, expect, it } from 'vitest';

import { markAllDirty, roundTrip } from './save-fidelity-harness';

const SLIDE = 'ppt/slides/slide1.xml';

async function warpAfterSave(warp: string): Promise<string> {
	const { saved } = await roundTrip('degenerate-shape.pptx', {
		patchParts: {
			[SLIDE]: (xml) =>
				xml.replace(
					'<a:bodyPr/><a:lstStyle/><a:p/>',
					`<a:bodyPr>${warp}</a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Warped</a:t></a:r></a:p>`,
				),
		},
		mutate: (data) => {
			markAllDirty(data);
		},
	});
	const xml = await saved.file(SLIDE)!.async('string');
	return /<a:prstTxWarp[\s\S]*?<\/a:prstTxWarp>/.exec(xml)?.[0] ?? '';
}

describe('prstTxWarp adjustment names', () => {
	it('keeps the adj1 and adj2 guides of a wave', async () => {
		const warp = await warpAfterSave(
			'<a:prstTxWarp prst="textWave4"><a:avLst><a:gd name="adj1" fmla="val 9000"/><a:gd name="adj2" fmla="val 2500"/></a:avLst></a:prstTxWarp>',
		);
		expect(warp).toContain('name="adj1" fmla="val 9000"');
		expect(warp).toContain('name="adj2" fmla="val 2500"');
	});

	it('keeps the single adj guide of an arch', async () => {
		const warp = await warpAfterSave(
			'<a:prstTxWarp prst="textArchUp"><a:avLst><a:gd name="adj" fmla="val 16200000"/></a:avLst></a:prstTxWarp>',
		);
		expect(warp).toContain('name="adj" fmla="val 16200000"');
	});
});
