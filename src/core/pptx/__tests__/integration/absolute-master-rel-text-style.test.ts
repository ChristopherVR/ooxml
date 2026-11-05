import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';

/**
 * `absolute-path-rels.pptx` writes every relationship target package-rooted
 * (`/ppt/slideMasters/slideMaster11.xml`). The layout-to-master lookup turned
 * that into `ppt//ppt/slideMasters/...`, so the master's `p:bodyStyle` never
 * reached the slide's body placeholders: no theme font (the text fell back to
 * the UI font) and no `spcBef`/`spcAft` between paragraphs. PowerPoint
 * measured 8 pt of spcAft plus 20 % spcBef between the bullets.
 */
describe('layout to master relationship with an absolute target', () => {
	it('inherits the master body style (font and paragraph spacing)', async () => {
		const buf = readFileSync(resolve(__dirname, '../fixtures/e2e/absolute-path-rels.pptx'));
		const data = await new PptxHandler().load(
			buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer,
		);
		const elements = data.slides[0]?.elements ?? [];
		const intro = elements.find(
			(e) => 'textSegments' in e && e.textSegments?.[0]?.text.startsWith('A TIME CAPSULE'),
		);
		expect(intro).toBeDefined();
		const seg = (intro as { textSegments: { style: { fontFamily?: string } }[] }).textSegments[0];
		expect(seg?.style.fontFamily).toBe('Mangal');
		const spacing = (
			intro as {
				textStyle: {
					inheritedParagraphSpacing: Record<
						string,
						{ paragraphSpacingBefore?: number; paragraphSpacingAfter?: number }
					>;
				};
			}
		).textStyle.inheritedParagraphSpacing['0'];
		expect(spacing?.paragraphSpacingAfter).toBe(8);
		expect(spacing?.paragraphSpacingBefore).toBeGreaterThan(0);
	});
});
