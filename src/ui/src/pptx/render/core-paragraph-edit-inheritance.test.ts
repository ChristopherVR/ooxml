import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { hasTextProperties, PptxHandler } from 'ooxml-core/pptx';
import { applyTextStyleUpdate } from './text-style-update';
import { resolveParagraphSpacing } from './paragraph-spacing';

it('projects a core-backed body edit over inherited spacing without mutating parsed snapshots', async () => {
	const source = await readFile(
		new URL('../../../../core/pptx/__tests__/fixtures/e2e/linked-textbox.pptx', import.meta.url),
	);
	const zip = await JSZip.loadAsync(source);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const body =
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:lnSpc><a:spcPts val="1286"/></a:lnSpc><a:spcAft><a:spcPts val="380"/></a:spcAft></a:pPr><a:r><a:t>First</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p><a:p><a:r><a:t>Third</a:t></a:r></a:p></p:txBody>';
	zip.file('ppt/slides/slide1.xml', xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/gu, body));
	const handler = new PptxHandler();
	const reopened = new PptxHandler();
	try {
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const target = data.slides[0]!.elements.find(
			(element) => hasTextProperties(element) && element.text?.startsWith('First'),
		);
		if (!target || !hasTextProperties(target)) throw new Error('Missing parsed text body');
		const before = structuredClone(target);
		const inherited = target.textStyle!.inheritedParagraphSpacing!;
		const baseline = target.textStyle!.resolvedParagraphGeometry!;
		expect(inherited[0]).toMatchObject({ lineSpacing: 1, paragraphSpacingAfter: 0 });
		expect(baseline.lineSpacingExactPt).toBe(12.86);
		const { patch } = applyTextStyleUpdate(
			target,
			{ lineSpacing: 2, paragraphSpacingAfter: 18 },
			null,
		);
		expect(target).toEqual(before);
		Object.assign(target, patch);
		expect(target.textStyle?.inheritedParagraphSpacing).toBe(inherited);
		expect(target.textStyle?.resolvedParagraphGeometry).toBe(baseline);
		expect(inherited).toEqual(before.textStyle?.inheritedParagraphSpacing);
		expect(baseline).toEqual(before.textStyle?.resolvedParagraphGeometry);
		for (const text of ['First', 'Second']) {
			const paragraph = target.textSegments!.find((segment) => segment.text === text)!;
			expect(
				resolveParagraphSpacing({
					paraProps: paragraph.paragraphProperties,
					bodyStyle: target.textStyle,
					paragraphLevel: paragraph.paragraphLevel ?? 0,
				}),
			).toMatchObject({ lineHeight: 2.4, spaceAfterPx: 18 });
		}
		const saved = await reopened.load(await handler.save(data.slides));
		const result = saved.slides[0]!.elements.find(
			(element) => hasTextProperties(element) && element.text?.startsWith('First'),
		);
		if (!result || !hasTextProperties(result)) throw new Error('Missing saved body edit');
		const second = result.textSegments!.find((segment) => segment.text === 'Second')!;
		expect(
			resolveParagraphSpacing({
				paraProps: second.paragraphProperties,
				bodyStyle: result.textStyle,
				paragraphLevel: second.paragraphLevel ?? 0,
			}),
		).toMatchObject({ lineHeight: 2.4, spaceAfterPx: 18 });
	} finally {
		handler.dispose();
		reopened.dispose();
	}
});
