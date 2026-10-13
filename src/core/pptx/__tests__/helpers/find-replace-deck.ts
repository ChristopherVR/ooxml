import JSZip from 'jszip';
import { PresentationBuilder } from 'ooxml-core/pptx';

const runs =
	'<a:r><a:rPr b="1" lang="en-US"/><a:t>Annual re</a:t></a:r>' +
	'<a:r><a:rPr i="1" u="sng"><a:hlinkClick r:id="rIdKeep"/></a:rPr><a:t>port</a:t></a:r>';
const txBody = `<a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r"/>${runs}<a:endParaRPr lang="fr-FR"/></a:p>`;
const box = (id: number) =>
	`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Rich ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
	`<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${txBody}</p:txBody></p:sp>`;

export async function createRichFindDeck(includeNotes = false) {
	const { handler, data } = await PresentationBuilder.create({ initialSlideCount: 1 });
	if (includeNotes) {
		data.slides[0]!.notes = 'Speaker report';
		data.slides[0]!.notesSegments = [
			{ text: 'Speaker re', style: { bold: true } },
			{ text: 'port', style: { italic: true } },
		];
	}
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const group =
		'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
		'<p:grpSpPr><a:xfrm><a:off x="0" y="1000000"/><a:ext cx="2000000" cy="500000"/><a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="500000"/></a:xfrm></p:grpSpPr>' +
		box(21) +
		'</p:grpSp>';
	const table =
		'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="30" name="Rich table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
		'<p:xfrm><a:off x="0" y="2000000"/><a:ext cx="2000000" cy="500000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
		'<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2000000"/></a:tblGrid><a:tr h="500000"><a:tc>' +
		`<a:txBody>${txBody}</a:txBody><a:tcPr><a:extLst><a:ext uri="opaque-cell"/></a:extLst></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
	const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	zip.file(
		'ppt/slides/slide1.xml',
		slideXml.replace('</p:spTree>', `${box(10)}${group}${table}</p:spTree>`),
	);
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
	zip.file(
		'ppt/slides/_rels/slide1.xml.rels',
		rels.replace(
			'</Relationships>',
			'<Relationship Id="rIdKeep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>',
		),
	);
	return zip.generateAsync({ type: 'arraybuffer' });
}
