/**
 * A deck built with think-cell must open with no compatibility notice.
 *
 * think-cell keeps its chart data in a hidden 1588-EMU OLE frame named "think-cell data - do not
 * delete" on every slide it touches, and PowerPoint writes every OLE frame as an
 * `mc:Choice Requires="v"` (VML) beside an `mc:Fallback` that holds the same `p:oleObj` plus its
 * preview picture. A 22-slide board deck reported both UNSUPPORTED_ALTERNATE_CONTENT_CHOICE and
 * PARTIAL_OLE_SUPPORT on every load for objects nobody can see. The frame below is copied from
 * that deck (ids and the creation GUID aside).
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';

/** A 1x1 transparent PNG, the preview picture of each OLE frame. */
const PNG_1X1 = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
	),
	(c) => c.charCodeAt(0),
);

const NS =
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** An OLE graphic frame the way PowerPoint writes it: VML choice, DrawingML fallback. */
function oleFrame(options: {
	id: number;
	name: string;
	progId: string;
	hidden: boolean;
	oleRel: string;
	imageRel: string;
	extent: number;
}): string {
	const { id, name, progId, hidden, oleRel, imageRel, extent } = options;
	const hiddenAttr = hidden ? ' hidden="1"' : '';
	const oleObj = (withPicture: boolean) =>
		`<p:oleObj name="${name}" r:id="${oleRel}" imgW="360" imgH="360" progId="${progId}"><p:embed/>` +
		(withPicture
			? `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"${hiddenAttr}/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
				`<p:blipFill><a:blip r:embed="${imageRel}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
				`<p:spPr><a:xfrm><a:off x="1588" y="1588"/><a:ext cx="${extent}" cy="${extent}"/></a:xfrm>` +
				`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
			: '') +
		`</p:oleObj>`;
	return (
		`<p:graphicFrame ${NS}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"${hiddenAttr}/>` +
		`<p:cNvGraphicFramePr><a:graphicFrameLocks/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
		`<p:xfrm><a:off x="1588" y="1588"/><a:ext cx="${extent}" cy="${extent}"/></p:xfrm>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">` +
		`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
		`<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">${oleObj(false)}</mc:Choice>` +
		`<mc:Fallback>${oleObj(true)}</mc:Fallback>` +
		`</mc:AlternateContent></a:graphicData></a:graphic></p:graphicFrame>`
	);
}

/** One blank slide plus the given OLE frames (each with its own embedding and preview part). */
async function deckWithOleFrames(
	frames: ReadonlyArray<{ name: string; progId: string; hidden: boolean; extent: number }>,
): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));

	const slidePath = 'ppt/slides/slide1.xml';
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	let slideXml = await zip.file(slidePath)!.async('string');
	let rels = await zip.file(relsPath)!.async('string');
	const xml = frames.map((frame, index) => {
		const n = index + 1;
		zip.file(`ppt/embeddings/oleObject${n}.bin`, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));
		zip.file(`ppt/media/olePreview${n}.png`, PNG_1X1);
		rels = rels.replace(
			'</Relationships>',
			`<Relationship Id="rIdOle${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject${n}.bin"/>` +
				`<Relationship Id="rIdOleImg${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/olePreview${n}.png"/></Relationships>`,
		);
		return oleFrame({ ...frame, id: 900 + n, oleRel: `rIdOle${n}`, imageRel: `rIdOleImg${n}` });
	});
	slideXml = slideXml.replace('</p:spTree>', `${xml.join('')}</p:spTree>`);
	zip.file(slidePath, slideXml);
	zip.file(relsPath, rels);

	let contentTypes = await zip.file('[Content_Types].xml')!.async('string');
	for (const [ext, type] of [
		['bin', 'application/vnd.openxmlformats-officedocument.oleObject'],
		['png', 'image/png'],
	]) {
		if (!contentTypes.includes(`Extension="${ext}"`)) {
			contentTypes = contentTypes.replace(
				'<Override',
				`<Default Extension="${ext}" ContentType="${type}"/><Override`,
			);
		}
	}
	zip.file('[Content_Types].xml', contentTypes);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

const THINK_CELL = {
	name: 'think-cell data - do not delete',
	progId: 'TCLayout.ActiveDocument.1',
	hidden: true,
	extent: 1588,
};

describe('think-cell deck compatibility notices', () => {
	it('opens a deck with a hidden think-cell OLE frame with no compatibility warning', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await deckWithOleFrames([THINK_CELL]));

		const ole = data.slides[0].elements.find((element) => element.type === 'ole');
		expect(ole).toMatchObject({ type: 'ole', name: THINK_CELL.name, hidden: true });
		expect(handler.getCompatibilityWarnings()).toStrictEqual([]);
	});

	it('still reports PARTIAL_OLE_SUPPORT for a visible embedded object', async () => {
		const handler = new PptxHandler();
		await handler.load(
			await deckWithOleFrames([
				THINK_CELL,
				{ name: 'Packager Shell Object', progId: 'Package', hidden: false, extent: 914400 },
			]),
		);

		const codes = handler.getCompatibilityWarnings().map((warning) => warning.code);
		expect(codes).toStrictEqual(['PARTIAL_OLE_SUPPORT']);
	});
});
