// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Picture watermarks: the VML picture shape (`WordPictureWatermark`) Word puts in a header
// paragraph, centred on the margins behind the text. The `_x0000_t75` picture-frame shapetype
// and the shape below match what Word 16 writes when it saves such a watermark; washout is
// `v:imagedata` gain 0.3 and blacklevel 0.35 in VML's 16.16 fixed-point form, which Word reads
// back as Brightness 85%, Contrast 15% and the Watermark colour type.
import type { InlineImage, PictureWatermarkSpec } from './inline-model';
import { makeNS, makeW, REL_NS, getR, type XmlDocument, type XmlElement } from './xml';

const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const WORD10_NS = 'urn:schemas-microsoft-com:office:word';

export const PICTURE_WATERMARK_PREFIX = 'WordPictureWatermark';
export const WASHOUT_GAIN = '19661f';
export const WASHOUT_BLACKLEVEL = '22938f';
/**
 * Word's washout as a linear channel map, x -> min(1, SLOPE x + INTERCEPT): black rises to light
 * grey and white stays white (measured from Word's PDF of a black and white picture). The PDF
 * painter, Print Layout's CSS filter and the model's luminance effect all derive from these.
 */
export const WASHOUT_SLOPE = 0.3;
export const WASHOUT_INTERCEPT = 0.8;
/** The same map as DrawingML luminance (the transfer is `contrast + 1` slope). */
export const WASHOUT_LUMINANCE = (() => {
	const half = (1 - WASHOUT_SLOPE) / 2;
	return { brightness: (WASHOUT_INTERCEPT - half) / (1 - half), contrast: WASHOUT_SLOPE - 1 };
})();
/** The same map as a CSS `contrast()` then `brightness()` filter. */
export const WASHOUT_CSS_FILTER = (() => {
	const contrast = WASHOUT_SLOPE / (WASHOUT_SLOPE + 2 * WASHOUT_INTERCEPT);
	const brightness = WASHOUT_SLOPE + 2 * WASHOUT_INTERCEPT;
	return `contrast(${Number(contrast.toFixed(4))}) brightness(${Number(brightness.toFixed(4))})`;
})();
/** Word's Scale choices in the Printed Watermark dialog, as percentages. */
export const PICTURE_WATERMARK_SCALES = [500, 200, 150, 100, 50] as const;

const FORMULAS = [
	'if lineDrawn pixelLineWidth 0',
	'sum @0 1 0',
	'sum 0 0 @1',
	'prod @2 1 2',
	'prod @3 21600 pixelWidth',
	'prod @3 21600 pixelHeight',
	'sum @0 0 1',
	'prod @6 1 2',
	'prod @7 21600 pixelWidth',
	'sum @8 21600 0',
	'prod @7 21600 pixelHeight',
	'sum @10 21600 0',
];

/** A VML fraction: `19661f` is 19661/65536; a plain number is the value itself. */
function fraction(value: string | null): number | undefined {
	const text = (value ?? '').trim();
	if (!text) return undefined;
	const number = Number(text.replace(/f$/i, ''));
	if (!Number.isFinite(number)) return undefined;
	return /f$/i.test(text) ? number / 65536 : number;
}

export function isVmlPictureWashout(data: XmlElement): boolean {
	const gain = fraction(data.getAttribute('gain'));
	const black = fraction(data.getAttribute('blacklevel'));
	return gain !== undefined && gain < 1 && black !== undefined && black > 0;
}

/**
 * The picture watermark a `v:shape` describes, or undefined when it is not one. `resolve` maps the
 * `v:imagedata` relationship to the header part's picture.
 */
export function parsePictureWatermarkShape(
	shape: XmlElement | undefined,
	resolve: (relId: string) => Pick<InlineImage, 'relId' | 'partName' | 'contentType'> | undefined,
): InlineImage | undefined {
	if (!shape || !(shape.getAttribute('id') ?? '').startsWith(PICTURE_WATERMARK_PREFIX))
		return undefined;
	const data = Array.from(shape.getElementsByTagNameNS(VML_NS, 'imagedata'))[0];
	const relId = data && getR(data, 'id');
	const picture = relId ? resolve(relId) : undefined;
	if (!data || !picture) return undefined;
	const style = shape.getAttribute('style') ?? '';
	const width = Number(/width:\s*([\d.]+)pt/i.exec(style)?.[1] ?? 0);
	const height = Number(/height:\s*([\d.]+)pt/i.exec(style)?.[1] ?? 0);
	const title = data.getAttributeNS(OFFICE_NS, 'title');
	const spec: PictureWatermarkSpec = {
		kind: 'picture',
		partName: picture.partName,
		contentType: picture.contentType,
		widthPt: width,
		heightPt: height,
		washout: isVmlPictureWashout(data),
		...(title ? { title } : {}),
	};
	// The run itself takes no room: the spec's size is where the picture is drawn behind the text.
	return {
		...picture,
		widthPx: 0,
		heightPx: 0,
		unsupported: 'Watermark',
		pictureWatermark: spec,
	};
}

/**
 * The drawn size for a picture of `natural` points: a percentage of it, or for `auto` the largest
 * size that fits the text area `box` with the picture's proportions.
 */
export function pictureWatermarkSizePt(
	natural: { widthPt: number; heightPt: number },
	scale: 'auto' | number,
	box: { widthPt: number; heightPt: number },
): { widthPt: number; heightPt: number } {
	const w = Math.max(1, natural.widthPt);
	const h = Math.max(1, natural.heightPt);
	const factor = scale === 'auto' ? Math.min(box.widthPt / w, box.heightPt / h) : scale / 100;
	const round = (value: number) => Math.round(value * 100) / 100;
	return { widthPt: round(w * factor), heightPt: round(h * factor) };
}

/** The `_x0000_t75` picture-frame `v:shapetype` Word writes before every VML picture shape. */
export function pictureFrameShapeType(doc: XmlDocument): XmlElement {
	const v = (name: string) => makeNS(doc, VML_NS, `v:${name}`);
	const type = v('shapetype');
	type.setAttribute('id', '_x0000_t75');
	type.setAttribute('coordsize', '21600,21600');
	type.setAttributeNS(OFFICE_NS, 'o:spt', '75');
	type.setAttributeNS(OFFICE_NS, 'o:preferrelative', 't');
	type.setAttribute('path', 'm@4@5l@4@11@9@11@9@5xe');
	type.setAttribute('filled', 'f');
	type.setAttribute('stroked', 'f');
	const stroke = v('stroke');
	stroke.setAttribute('joinstyle', 'miter');
	const formulas = v('formulas');
	for (const eqn of FORMULAS) {
		const f = v('f');
		f.setAttribute('eqn', eqn);
		formulas.appendChild(f);
	}
	const path = v('path');
	path.setAttributeNS(OFFICE_NS, 'o:extrusionok', 'f');
	path.setAttribute('gradientshapeok', 't');
	path.setAttributeNS(OFFICE_NS, 'o:connecttype', 'rect');
	const lock = makeNS(doc, OFFICE_NS, 'o:lock');
	lock.setAttributeNS(VML_NS, 'v:ext', 'edit');
	lock.setAttribute('aspectratio', 't');
	for (const child of [stroke, formulas, path, lock]) type.appendChild(child);
	return type;
}

/** A `w:r` holding Word's picture-watermark VML for `spec`, pointing at relationship `relId`. */
export function buildPictureWatermarkRun(
	doc: XmlDocument,
	spec: PictureWatermarkSpec,
	relId: string,
	uniqueId: string,
): XmlElement {
	const v = (name: string) => makeNS(doc, VML_NS, `v:${name}`);
	const run = makeW(doc, 'r');
	const props = makeW(doc, 'rPr');
	props.appendChild(makeW(doc, 'noProof'));
	run.appendChild(props);
	const pict = makeW(doc, 'pict');
	const type = pictureFrameShapeType(doc);
	const shape = v('shape');
	shape.setAttribute('id', `${PICTURE_WATERMARK_PREFIX}${uniqueId}`);
	shape.setAttributeNS(OFFICE_NS, 'o:spid', `_x0000_s${2048 + (Number(uniqueId) || 1)}`);
	shape.setAttribute('type', '#_x0000_t75');
	shape.setAttribute(
		'style',
		`position:absolute;margin-left:0;margin-top:0;width:${spec.widthPt}pt;height:${spec.heightPt}pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin`,
	);
	shape.setAttributeNS(OFFICE_NS, 'o:allowincell', 'f');
	const data = v('imagedata');
	data.setAttributeNS(REL_NS, 'r:id', relId);
	data.setAttributeNS(OFFICE_NS, 'o:title', spec.title ?? '');
	if (spec.washout) {
		data.setAttribute('gain', WASHOUT_GAIN);
		data.setAttribute('blacklevel', WASHOUT_BLACKLEVEL);
	}
	const wrap = makeNS(doc, WORD10_NS, 'w10:wrap');
	wrap.setAttribute('anchorx', 'margin');
	wrap.setAttribute('anchory', 'margin');
	shape.appendChild(data);
	shape.appendChild(wrap);
	pict.appendChild(type);
	pict.appendChild(shape);
	run.appendChild(pict);
	return run;
}
