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

it('authors whole-body spacing even when the chosen value matches the first paragraph baseline', async () => {
	const source = await readFile(
		new URL('../../../../core/pptx/__tests__/fixtures/e2e/linked-textbox.pptx', import.meta.url),
	);
	const zip = await JSZip.loadAsync(source);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const body =
		'<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:lvl1pPr></a:lstStyle><a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="2400"/><a:t>First</a:t></a:r></a:p><a:p><a:r><a:rPr sz="2400"/><a:t>Second</a:t></a:r></a:p></p:txBody>';
	zip.file('ppt/slides/slide1.xml', xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/gu, body));
	const handler = new PptxHandler();
	const reopened = new PptxHandler();
	try {
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const target = data.slides[0]!.elements.find(
			(e) => hasTextProperties(e) && e.text?.startsWith('First'),
		)!;
		if (!hasTextProperties(target)) throw new Error('Missing text');
		const before = structuredClone(target);
		const inherited = target.textStyle!.inheritedParagraphSpacing!;
		const baseline = target.textStyle!.resolvedParagraphGeometry!;
		expect(target.textStyle!.lineSpacing).toBe(1.2);
		expect(inherited[0]!.lineSpacing).toBe(1.5);
		const secondBefore = target.textSegments!.find((s) => s.text === 'Second')!;
		expect(secondBefore.paragraphProperties?.lineSpacing).toBeUndefined();
		expect(
			resolveParagraphSpacing({
				paraProps: secondBefore.paragraphProperties,
				bodyStyle: target.textStyle,
			}).lineHeight,
		).toBeCloseTo(1.8);
		const { patch } = applyTextStyleUpdate(target, { lineSpacing: 1.2 }, null);
		expect(target).toEqual(before);
		Object.assign(target, patch);
		expect(target.textStyle!.inheritedParagraphSpacing).toBe(inherited);
		expect(target.textStyle!.resolvedParagraphGeometry).toBe(baseline);
		expect(inherited).toEqual(before.textStyle!.inheritedParagraphSpacing);
		expect(baseline).toEqual(before.textStyle!.resolvedParagraphGeometry);
		const second = target.textSegments!.find((s) => s.text === 'Second')!;
		expect(
			resolveParagraphSpacing({
				paraProps: second.paragraphProperties,
				bodyStyle: target.textStyle,
			}).lineHeight,
		).toBeCloseTo(1.44);
		const saved = await reopened.load(await handler.save(data.slides));
		const result = saved.slides[0]!.elements.find(
			(e) => hasTextProperties(e) && e.text?.startsWith('First'),
		)!;
		if (!hasTextProperties(result)) throw new Error('Missing reopened text');
		const reopenedSecond = result.textSegments!.find((s) => s.text === 'Second')!;
		expect(
			resolveParagraphSpacing({
				paraProps: reopenedSecond.paragraphProperties,
				bodyStyle: result.textStyle,
			}).lineHeight,
		).toBeCloseTo(1.44);
		expect(result.textStyle!.inheritedParagraphSpacing![0]!.lineSpacing).toBe(1.5);
		Object.assign(target, before);
		const undone = await reopened.load(await handler.save(data.slides));
		const restored = undone.slides[0]!.elements.find(
			(e) => hasTextProperties(e) && e.text?.startsWith('First'),
		)!;
		if (!hasTextProperties(restored)) throw new Error('Missing restored text');
		const restoredSecond = restored.textSegments!.find((s) => s.text === 'Second')!;
		expect(restoredSecond.paragraphProperties?.lineSpacing).toBeUndefined();
		expect(
			resolveParagraphSpacing({
				paraProps: restoredSecond.paragraphProperties,
				bodyStyle: restored.textStyle,
			}).lineHeight,
		).toBeCloseTo(1.8);
	} finally {
		handler.dispose();
		reopened.dispose();
	}
});
