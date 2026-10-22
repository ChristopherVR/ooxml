/**
 * Regression guard for issue #140: a title placeholder whose layout counterpart
 * paints a picture fill (`p:spPr/a:blipFill`) while its own `spPr` is empty got
 * `fillMode: 'image'` from inheritance but never an image, so the grey panel
 * behind the title was missing. The blip's relationship id lives in the
 * LAYOUT's rels, not the slide's.
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxElement } from '../../types';

const NS =
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const GROUP_PROPS =
	'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS} type="title"><p:cSld><p:spTree>${GROUP_PROPS}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="1000000"/><a:ext cx="5000000" cy="1200000"/></a:xfrm><a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`;

const SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS}><p:cSld><p:spTree>${GROUP_PROPS}<p:sp><p:nvSpPr><p:cNvPr id="3" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Panel title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;

const SLIDE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`;

const PNG_1X1 = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
	),
	(c) => c.charCodeAt(0),
);

const fixture = fileURLToPath(
	new URL('../../../__tests__/fixtures/e2e/linked-textbox.pptx', import.meta.url),
);

async function loadSlide(): Promise<PptxElement[]> {
	const zip = await JSZip.loadAsync(readFileSync(fixture));
	const layoutRelsPath = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels';
	const layoutRels = await zip.file(layoutRelsPath)!.async('string');
	zip.file(
		layoutRelsPath,
		layoutRels.replace(
			'</Relationships>',
			'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layoutPanel.png"/></Relationships>',
		),
	);
	zip.file('ppt/slideLayouts/slideLayout1.xml', LAYOUT_XML);
	zip.file('ppt/media/layoutPanel.png', PNG_1X1);
	zip.file('ppt/slides/slide1.xml', SLIDE_XML);
	zip.file('ppt/slides/_rels/slide1.xml.rels', SLIDE_RELS_XML);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	const data = await new PptxHandler().load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	return data.slides[0]!.elements;
}

describe('a placeholder inheriting a picture fill from its layout', () => {
	it('resolves a displayable fillImageUrl and paints as a shape', async () => {
		const elements = await loadSlide();
		const title = elements.find((el) => 'name' in el && el.name === 'Title 1');
		const style = (title as { shapeStyle?: Record<string, unknown> } | undefined)?.shapeStyle;
		expect(style?.fillMode).toBe('image');
		expect(String(style?.fillImageUrl)).toMatch(/^(?:data:image|blob:)/u);
		expect(title?.type).toBe('shape');
	});
});
