import JSZip from 'jszip';
import { RELATIONSHIP_TYPES } from '../../opc/index';
/** Synthetic cached picture tree over native Word/Excel packages, with crop and ordered effects. */
export async function withDiagramPicture(
	bytes: Uint8Array,
	prefix: 'word' | 'xl',
): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	const drawing = `${prefix}/diagrams/drawing1.xml`;
	zip.file(
		drawing,
		`<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dsp:spTree><dsp:pic><dsp:nvPicPr><dsp:cNvPr id="1" name="Picture"/></dsp:nvPicPr><dsp:blipFill><a:blip r:embed="picture"><a:grayscl/><a:alphaModFix amt="50000"/></a:blip><a:srcRect l="25000" r="25000"/><a:stretch><a:fillRect/></a:stretch></dsp:blipFill><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></dsp:spPr></dsp:pic></dsp:spTree></dsp:drawing>`,
	);
	zip.file(
		`${prefix}/diagrams/_rels/drawing1.xml.rels`,
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="picture" Type="${RELATIONSHIP_TYPES.image}" Target="../media/shared-picture.png"/><Relationship Id="external" Type="${RELATIONSHIP_TYPES.image}" Target="https://example.invalid/private.png" TargetMode="External"/></Relationships>`,
	);
	zip.file(
		`${prefix}/media/shared-picture.png`,
		'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP8zwACTGCSAQANHQEDgslx/wAAAABJRU5ErkJggg==',
		{ base64: true },
	);
	return zip.generateAsync({ type: 'uint8array' });
}
