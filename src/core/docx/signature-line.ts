// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Insert > Signature Line: the VML picture shape Word writes for a signature line
// (`o:signatureline` inside a `v:shape`), with a placeholder picture. Only the line is inserted;
// signing it needs a certificate and a signing service, which this package does not have.
import { encodePng } from '../drawingml/raster-encode';
import type { InlineImage } from './inline-model';
import type { Block } from './model';
import type { Paragraph } from './model-paragraph';
import { pictureFrameShapeType } from './picture-watermark';
import { makeNS, makeW, REL_NS, type XmlDocument, type XmlElement } from './xml';

const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';

/** What the Signature Setup dialog collects. */
export interface SignatureLineSpec {
	/** `{GUID}` identifying the line; a signature refers to it. */
	id: string;
	suggestedSigner?: string;
	suggestedSignerTitle?: string;
	suggestedSignerEmail?: string;
	instructions?: string;
	/** The signer may add comments with the signature (default on). */
	allowComments: boolean;
	/** The signing date is shown beside the signature (default on). */
	showSignDate: boolean;
}

/** The size Word gives a new signature line: 192 by 96 points. */
export const SIGNATURE_LINE_SIZE_PT = { width: 192, height: 96 } as const;

const flag = (value: string | null): boolean => value === 't' || value === 'true' || value === '1';
const text = (value: string | null): string | undefined => (value ? value : undefined);

/** The signature line a `v:shape` carries, or undefined when it has none. */
export function parseSignatureLineShape(
	shape: XmlElement | undefined,
): SignatureLineSpec | undefined {
	const line = shape && Array.from(shape.getElementsByTagNameNS(OFFICE_NS, 'signatureline'))[0];
	if (!line) return undefined;
	const office = (name: string) => text(line.getAttributeNS(OFFICE_NS, name));
	const set =
		line.getAttribute('signinginstructionsset') ??
		line.getAttributeNS(OFFICE_NS, 'signinginstructionsset');
	const instructions = flag(set) ? office('signinginstructions') : undefined;
	const signer = office('suggestedsigner');
	const title = office('suggestedsigner2');
	const email = office('suggestedsigneremail');
	const comments = line.getAttribute('allowcomments');
	const signDate = line.getAttribute('showsigndate');
	return {
		id: line.getAttribute('id') ?? '',
		...(signer ? { suggestedSigner: signer } : {}),
		...(title ? { suggestedSignerTitle: title } : {}),
		...(email ? { suggestedSignerEmail: email } : {}),
		...(instructions ? { instructions } : {}),
		// CT_SignatureLine: a missing showsigndate is true and a missing allowcomments is false.
		allowComments: flag(comments),
		showSignDate: signDate === null || flag(signDate),
	};
}

/** A new signature line's picture reference and size, for the run model. */
export function signatureLineImage(
	spec: SignatureLineSpec,
	picture: Pick<InlineImage, 'partName' | 'contentType'>,
	altText = 'Signature line',
): InlineImage {
	return {
		relId: '',
		partName: picture.partName,
		contentType: picture.contentType,
		widthPx: Math.round((SIGNATURE_LINE_SIZE_PT.width * 96) / 72),
		heightPx: Math.round((SIGNATURE_LINE_SIZE_PT.height * 96) / 72),
		altText,
		signatureLine: spec,
	};
}

/** A `w:r` holding a signature line for `spec`, its picture at relationship `relId`. */
export function buildSignatureLineRun(
	doc: XmlDocument,
	spec: SignatureLineSpec,
	relId: string,
	uniqueId: string,
	size: { widthPt: number; heightPt: number },
): XmlElement {
	const v = (name: string) => makeNS(doc, VML_NS, `v:${name}`);
	const run = makeW(doc, 'r');
	const props = makeW(doc, 'rPr');
	props.appendChild(makeW(doc, 'noProof'));
	run.appendChild(props);
	const pict = makeW(doc, 'pict');
	const shape = v('shape');
	shape.setAttribute('id', `_x0000_i${1024 + (Number(uniqueId) || 1)}`);
	shape.setAttribute('type', '#_x0000_t75');
	shape.setAttribute('alt', 'Signature line');
	shape.setAttribute('style', `width:${size.widthPt}pt;height:${size.heightPt}pt`);
	const data = v('imagedata');
	data.setAttributeNS(REL_NS, 'r:id', relId);
	data.setAttributeNS(OFFICE_NS, 'o:title', '');
	const lock = makeNS(doc, OFFICE_NS, 'o:lock');
	lock.setAttributeNS(VML_NS, 'v:ext', 'edit');
	for (const name of ['ungrouping', 'rotation', 'cropping', 'verticies'])
		lock.setAttribute(name, 't');
	const line = makeNS(doc, OFFICE_NS, 'o:signatureline');
	line.setAttributeNS(VML_NS, 'v:ext', 'edit');
	line.setAttribute('id', spec.id);
	line.setAttribute('provid', '{00000000-0000-0000-0000-000000000000}');
	const set = (name: string, value: string | undefined) => {
		if (value) line.setAttributeNS(OFFICE_NS, `o:${name}`, value);
	};
	set('suggestedsigner', spec.suggestedSigner);
	set('suggestedsigner2', spec.suggestedSignerTitle);
	set('suggestedsigneremail', spec.suggestedSignerEmail);
	if (spec.instructions) {
		line.setAttribute('signinginstructionsset', 't');
		line.setAttributeNS(OFFICE_NS, 'o:signinginstructions', spec.instructions);
	}
	line.setAttribute('showsigndate', spec.showSignDate ? 't' : 'f');
	line.setAttribute('allowcomments', spec.allowComments ? 't' : 'f');
	line.setAttribute('issignatureline', 't');
	for (const child of [data, lock, line]) shape.appendChild(child);
	pict.appendChild(pictureFrameShapeType(doc));
	pict.appendChild(shape);
	run.appendChild(pict);
	return run;
}

/**
 * The placeholder picture of a new signature line: a cross at the left and the signing line under
 * it, drawn here (not copied from any product) as a transparent PNG at twice the size in points.
 */
export function signatureLinePlaceholderPng(): Uint8Array {
	const width = SIGNATURE_LINE_SIZE_PT.width * 2;
	const height = SIGNATURE_LINE_SIZE_PT.height * 2;
	const rgb = new Uint8Array(width * height * 3);
	const alpha = new Uint8Array(width * height);
	const ink = (x: number, y: number) => {
		if (x < 0 || y < 0 || x >= width || y >= height) return;
		const at = y * width + x;
		alpha[at] = 255;
		rgb.fill(90, at * 3, at * 3 + 3);
	};
	const baseline = Math.round(height * 0.72);
	for (let x = 8; x < width - 8; x++) for (let t = 0; t < 2; t++) ink(x, baseline + t);
	for (let i = 0; i < 28; i++)
		for (let t = 0; t < 3; t++) {
			ink(14 + i + t, baseline - 40 + i);
			ink(14 + 28 - i + t, baseline - 40 + i);
		}
	return encodePng(width, height, rgb, alpha);
}

/** Every signature line in the body (and its tables), in document order. */
export function signatureLinesOf(blocks: readonly Block[]): SignatureLineSpec[] {
	const lines: SignatureLineSpec[] = [];
	const visit = (paragraph: Paragraph) => {
		for (const run of paragraph.runs)
			if (run.image?.signatureLine) lines.push(run.image.signatureLine);
	};
	for (const block of blocks) {
		if (block.type === 'paragraph') visit(block);
		else for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(visit);
	}
	return lines;
}
