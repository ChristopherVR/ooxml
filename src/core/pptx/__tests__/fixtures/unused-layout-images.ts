import JSZip from 'jszip';

import { buildBlankPresentationArchive } from 'ooxml-core/pptx';

export const USED_LAYOUT_PATH = 'ppt/slideLayouts/slideLayout1.xml';
export const UNUSED_LAYOUT_PATH = 'ppt/slideLayouts/slideLayout2.xml';

// Valid, distinct one-pixel PNGs, stretched into visible layout artwork.
const USED_PNG =
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGPQSd8HAAIUAVLOJNSwAAAAAElFTkSuQmCC';
const UNUSED_PNG =
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGN4EKkJAAN/AWPDY9tWAAAAAElFTkSuQmCC';

function picture(id: number, description: string, width: number, height: number): string {
	return (
		`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${description}" descr="${description}"/>` +
		'<p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
		'<p:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
		`<p:spPr><a:xfrm><a:off x="9525000" y="381000"/><a:ext cx="${width * 9525}" cy="${height * 9525}"/></a:xfrm>` +
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
	);
}

/** One loaded slide uses layout 1; layout 2 has different embedded artwork but no slide. */
export async function buildUnusedLayoutImagesDeck(): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(
		await buildBlankPresentationArchive({ initialSlideCount: 1, title: 'Layout artwork' }),
	);
	// Keep only the two layouts exercised by the fixture.
	for (const path of Object.keys(zip.files)) {
		const index = /^ppt\/slideLayouts\/(?:_rels\/)?slideLayout(\d+)\.xml(?:\.rels)?$/u.exec(
			path,
		)?.[1];
		if (index && Number(index) > 2) zip.remove(path);
	}
	const masterPath = 'ppt/slideMasters/slideMaster1.xml';
	const master = await zip.file(masterPath)!.async('string');
	zip.file(
		masterPath,
		master.replace(/<p:sldLayoutId\b[^>]*\/>/gu, (node) =>
			/r:id="rId[12]"/u.test(node) ? node : '',
		),
	);
	const masterRelsPath = 'ppt/slideMasters/_rels/slideMaster1.xml.rels';
	const masterRels = await zip.file(masterRelsPath)!.async('string');
	zip.file(
		masterRelsPath,
		masterRels.replace(/<Relationship\b[^>]*\/>/gu, (node) =>
			node.includes('/slideLayout"') && !/slideLayout[12]\.xml/u.test(node) ? '' : node,
		),
	);
	const slideRelsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const slideRels = await zip.file(slideRelsPath)!.async('string');
	zip.file(slideRelsPath, slideRels.replace('slideLayout7.xml', 'slideLayout1.xml'));

	for (const [path, name, description, mediaName, png] of [
		[USED_LAYOUT_PATH, 'Used Logo Layout', 'Used Layout Logo', 'used-logo.png', USED_PNG],
		[UNUSED_LAYOUT_PATH, 'Unused Logo Layout', 'Unused Layout Logo', 'unused-logo.png', UNUSED_PNG],
	]) {
		const xml = await zip.file(path)!.async('string');
		const small =
			path === UNUSED_LAYOUT_PATH
				? picture(91, 'Small Layout Logo', 11, 13).replace('x="9525000"', 'x="11620500"')
				: '';
		zip.file(
			path,
			xml
				.replace(/<p:cSld\b[^>]*>/u, `<p:cSld name="${name}">`)
				.replace('</p:spTree>', `${picture(90, description, 180, 90)}${small}</p:spTree>`),
		);
		const relsPath = path.replace('slideLayouts/', 'slideLayouts/_rels/') + '.rels';
		const rels = await zip.file(relsPath)!.async('string');
		zip.file(
			relsPath,
			rels.replace(
				'</Relationships>',
				`<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/></Relationships>`,
			),
		);
		zip.file(`ppt/media/${mediaName}`, png, { base64: true });
	}

	const slidePath = 'ppt/slides/slide1.xml';
	const slide = await zip.file(slidePath)!.async('string');
	const title =
		'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Fixture Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
		'<p:spPr><a:xfrm><a:off x="476250" y="1714500"/><a:ext cx="7620000" cy="952500"/></a:xfrm>' +
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="3200"/>' +
		'<a:t>Layout artwork</a:t></a:r></a:p></p:txBody></p:sp>';
	zip.file(slidePath, slide.replace('</p:spTree>', `${title}</p:spTree>`));
	const types = (await zip.file('[Content_Types].xml')!.async('string')).replace(
		/<Override\b[^>]*\/>/gu,
		(node) =>
			node.includes('/ppt/slideLayouts/') && !/slideLayout[12]\.xml/u.test(node) ? '' : node,
	);
	zip.file(
		'[Content_Types].xml',
		types.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'),
	);
	return zip.generateAsync({ type: 'uint8array' });
}
