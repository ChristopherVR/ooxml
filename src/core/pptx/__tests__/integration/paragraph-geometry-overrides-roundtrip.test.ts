import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import { hasTextProperties } from '../../core/types';
import { resolveParagraphGeometryOverrides } from '../../editor/render/paragraph-geometry-overrides';

type ShapeXml = { 'p:txBody': { 'a:p': Record<string, unknown>[] } };

it('uses own paragraph geometry without changing source or raw paragraph XML', async () => {
	const source = await readFile(new URL('../fixtures/e2e/linked-textbox.pptx', import.meta.url));
	const zip = await JSZip.loadAsync(source);
	const sourceXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const body =
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr eaLnBrk="0" latinLnBrk="1" hangingPunct="0" fontAlgn="t" defTabSz="457200"><a:extLst><a:ext uri="{geometry-probe}"><g:retain xmlns:g="urn:geometry-probe" value="untouched"/></a:ext></a:extLst></a:pPr><a:r><a:t>First geometry</a:t></a:r></a:p><a:p><a:pPr eaLnBrk="1" fontAlgn="b"/><a:r><a:t>Second geometry</a:t></a:r></a:p></p:txBody>';
	const authoredXml = sourceXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/u, body);
	zip.file('ppt/slides/slide1.xml', authoredXml);
	const handler = new PptxHandler();
	const reopened = new PptxHandler();
	try {
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const element = data.slides[0]!.elements.find(
			(e) => hasTextProperties(e) && e.text?.startsWith('First geometry'),
		)!;
		if (!hasTextProperties(element)) throw new Error('Missing parsed geometry');
		const before = structuredClone(element);
		const first = element.textSegments!.find((s) => s.text === 'First geometry')!;
		const second = element.textSegments!.find((s) => s.text === 'Second geometry')!;
		expect(resolveParagraphGeometryOverrides(first.paragraphProperties, element.textStyle)).toEqual(
			{
				eaLineBreak: false,
				latinLineBreak: true,
				hangingPunctuation: false,
				fontAlignment: 't',
				defaultTabSize: 48,
			},
		);
		expect(
			resolveParagraphGeometryOverrides(second.paragraphProperties, element.textStyle),
		).toEqual({
			eaLineBreak: true,
			latinLineBreak: false,
			hangingPunctuation: true,
			fontAlignment: 'b',
			defaultTabSize: 96,
		});
		expect(element).toEqual(before);
		const saved = await handler.save(data.slides);
		const savedZip = await JSZip.loadAsync(saved);
		const parser = new XMLParser({ ignoreAttributes: false });
		const ownProperties = (xml: string) => {
			const shapes = parser.parse(xml)['p:sld']['p:cSld']['p:spTree']['p:sp'] as ShapeXml[];
			const shape = shapes.find((s) => JSON.stringify(s['p:txBody']).includes('First geometry'));
			if (!shape) throw new Error('Missing source XML shape');
			return shape['p:txBody']['a:p'].map((p) => p['a:pPr']);
		};
		expect(ownProperties(await savedZip.file('ppt/slides/slide1.xml')!.async('string'))).toEqual(
			ownProperties(authoredXml),
		);
		const result = await reopened.load(saved);
		const restored = result.slides[0]!.elements.find(
			(e) => hasTextProperties(e) && e.text?.startsWith('First geometry'),
		)!;
		if (!hasTextProperties(restored)) throw new Error('Missing reopened geometry');
		const restoredSecond = restored.textSegments!.find((s) => s.text === 'Second geometry')!;
		expect(
			resolveParagraphGeometryOverrides(restoredSecond.paragraphProperties, restored.textStyle),
		).toEqual(resolveParagraphGeometryOverrides(second.paragraphProperties, element.textStyle));
	} finally {
		handler.dispose();
		reopened.dispose();
	}
});
