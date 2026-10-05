import JSZip from 'jszip';
import { createVsdxFixture } from './fixture.mjs';
import { emf, record } from '../scripts/emf-audit/fixtures.mjs';
/** Original generated EMF primitives in a generated VSDX package. No corpus bytes. */
export async function createMetafileFixture(unsupported = false) {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Embedded vector'));
	const path = 'visio/pages/page1.xml';
	const page = await zip.file(path).async('string');
	zip.file(
		path,
		page.replace(
			'<Shapes>',
			'<Shapes><Shape ID="emf" Type="Foreign"><Cell N="Width" V="2"/><Cell N="Height" V="2"/><Cell N="PinX" V="2"/><Cell N="PinY" V="2"/><ForeignData ForeignType="EnhMetaFile"><Rel xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="emf"/></ForeignData></Shape>',
		),
	);
	zip.file(
		'visio/pages/_rels/page1.xml.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="emf" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/generated.emf"/></Relationships>',
	);
	zip.file(
		'visio/media/generated.emf',
		emf(unsupported ? [record(33), record(34, [-1])] : [record(43, [5, 10, 95, 90])], 100, 100),
	);
	return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
