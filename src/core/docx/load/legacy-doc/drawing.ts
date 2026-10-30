// Pictures, floating pictures and text boxes: DrawingML inside `w:drawing`. The .doc picture
// bytes become package media; floating shapes keep their anchor, position and wrapping.
import type {
	DocBlip,
	DocCrop,
	DocFloatingShape,
	DocInlinePicture,
} from '@christophervr/ole2/doc-objects-types';
import type { DocModelTextBox } from '@christophervr/ole2/doc-model-types';
import type { Conversion } from './context';
import { attrs, EMU_PER_TWIP, el, hex6, NS } from './xml';

/** What the drawing writers need from the block writer (text box content). */
export interface StoryWriter {
	readonly conversion: Conversion;
	storyXml(story: DocModelTextBox['content']): string;
}

const PIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const emu = (twips: number): number => Math.max(1, Math.round(twips * EMU_PER_TWIP));

function srcRect(crop: DocCrop | undefined): string {
	if (!crop) return '';
	const part = (value: number) => Math.round(value * 100000);
	if (!crop.top && !crop.bottom && !crop.left && !crop.right) return '';
	return el('a:srcRect', {
		l: part(crop.left) || undefined,
		t: part(crop.top) || undefined,
		r: part(crop.right) || undefined,
		b: part(crop.bottom) || undefined,
	});
}

function pictureGraphic(
	conversion: Conversion,
	blip: DocBlip,
	crop: DocCrop | undefined,
	cx: number,
	cy: number,
	name: string,
): string | undefined {
	const id = conversion.imageRelationship(blip);
	if (!id) {
		conversion.warn(
			`A ${blip.format.toUpperCase()} picture cannot be written to a Word document and was left out.`,
		);
		return undefined;
	}
	return (
		`<a:graphic><a:graphicData uri="${PIC_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0"${attrs({ name })}/><pic:cNvPicPr/></pic:nvPicPr>` +
		`<pic:blipFill><a:blip r:embed="${id}"/>${srcRect(crop)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
		`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`
	);
}

/** An inline picture (or an OLE object's preview) as `<w:drawing>`. */
export function inlinePictureXml(
	conversion: Conversion,
	picture: DocInlinePicture,
	fallbackName = 'Picture',
): string {
	if (!picture.blip) {
		conversion.warn(
			'A picture without image data (for example a linked INCLUDEPICTURE) was left out.',
		);
		return '';
	}
	const cx = emu(picture.widthTwips);
	const cy = emu(picture.heightTwips);
	const docPr = conversion.nextDocPrId();
	const name = picture.name || `${fallbackName} ${docPr}`;
	const graphic = pictureGraphic(conversion, picture.blip, picture.crop, cx, cy, name);
	if (!graphic) return '';
	return (
		`<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
		`<wp:docPr id="${docPr}"${attrs({ name, descr: picture.alt })}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:inline></w:drawing>`
	);
}

const RELATIVE_H: Record<DocFloatingShape['relativeX'], string> = {
	margin: 'margin',
	page: 'page',
	column: 'column',
	paragraph: 'column',
	unknown: 'column',
};
const RELATIVE_V: Record<DocFloatingShape['relativeY'], string> = {
	margin: 'margin',
	page: 'page',
	column: 'paragraph',
	paragraph: 'paragraph',
	unknown: 'paragraph',
};
const WRAP_SIDE = {
	both: 'bothSides',
	left: 'left',
	right: 'right',
	largest: 'largest',
	unknown: 'bothSides',
};

const TIGHT_POLYGON =
	'<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';

function wrapXml(shape: DocFloatingShape): string {
	const side = WRAP_SIDE[shape.wrapSide];
	switch (shape.wrap) {
		case 'topAndBottom':
			return '<wp:wrapTopAndBottom/>';
		case 'square':
			return `<wp:wrapSquare wrapText="${side}"/>`;
		case 'tight':
			return `<wp:wrapTight wrapText="${side}">${TIGHT_POLYGON}</wp:wrapTight>`;
		case 'through':
			return `<wp:wrapThrough wrapText="${side}">${TIGHT_POLYGON}</wp:wrapThrough>`;
		default:
			return '<wp:wrapNone/>';
	}
}

function anchorWrapper(
	conversion: Conversion,
	shape: DocFloatingShape,
	name: string,
	graphic: string,
): string {
	const cx = emu(shape.widthTwips);
	const cy = emu(shape.heightTwips);
	const docPr = conversion.nextDocPrId();
	return (
		`<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${conversion.nextZOrder()}" behindDoc="${shape.behindText ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
		`<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${RELATIVE_H[shape.relativeX]}"><wp:posOffset>${Math.round(shape.leftTwips * EMU_PER_TWIP)}</wp:posOffset></wp:positionH>` +
		`<wp:positionV relativeFrom="${RELATIVE_V[shape.relativeY]}"><wp:posOffset>${Math.round(shape.topTwips * EMU_PER_TWIP)}</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml(shape)}` +
		`<wp:docPr id="${docPr}"${attrs({ name, descr: shape.alt })}/><wp:cNvGraphicFramePr/>${graphic}</wp:anchor>`
	);
}

/** A floating picture as `<w:drawing><wp:anchor>`. */
export function floatingPictureXml(conversion: Conversion, shape: DocFloatingShape): string {
	if (!shape.blip) {
		conversion.warn('A floating picture whose image data could not be read was left out.');
		return '';
	}
	const name = shape.name || `Picture ${conversion.nextDocPrId()}`;
	const graphic = pictureGraphic(
		conversion,
		shape.blip,
		shape.crop,
		emu(shape.widthTwips),
		emu(shape.heightTwips),
		name,
	);
	return graphic ? `<w:drawing>${anchorWrapper(conversion, shape, name, graphic)}</w:drawing>` : '';
}

/** A floating text box as a `wps` shape whose content is the converted text box story. */
export function textBoxXml(
	writer: StoryWriter,
	shape: DocFloatingShape,
	textBox: DocModelTextBox | undefined,
): string {
	const { conversion } = writer;
	if (!textBox) {
		conversion.warn('A text box whose content could not be found was left out.');
		return '';
	}
	const cx = emu(shape.widthTwips);
	const cy = emu(shape.heightTwips);
	const fill =
		shape.fill?.enabled && shape.fill.colorRgb !== undefined
			? `<a:solidFill><a:srgbClr val="${hex6(shape.fill.colorRgb)}"/></a:solidFill>`
			: '<a:noFill/>';
	const line =
		shape.line?.enabled !== false && shape.line?.colorRgb !== undefined
			? `<a:ln w="${shape.line?.widthEmu ?? 9525}"><a:solidFill><a:srgbClr val="${hex6(shape.line.colorRgb)}"/></a:solidFill></a:ln>`
			: '<a:ln><a:noFill/></a:ln>';
	const graphic =
		`<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
		`<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}${line}</wps:spPr>` +
		`<wps:txbx><w:txbxContent>${writer.storyXml(textBox.content)}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>`;
	const name = shape.name || `Text Box ${textBox.id + 1}`;
	return `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="wps"><w:drawing>${anchorWrapper(conversion, shape, name, graphic)}</w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent>`;
}
