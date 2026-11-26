import {
	SIGNATURE_LINE_SIZE_PT,
	signatureLineImage,
	signatureLinePlaceholderPng,
	webVideoImage,
	webVideoPosterPng,
	type InlineImage,
	type SignatureLineSpec,
	type WebVideoSpec,
} from 'ooxml-core/docx';
import type { EditorView } from 'prosemirror-view';
import { focusView } from './focus-view';
import { schema } from './schema';

/** Stores generated picture bytes as a new media part and returns its part name. */
export type StageGenerated = (bytes: Uint8Array, contentType: string) => string;

const PNG = 'image/png';

/** Puts the picture at the selection as an inline image carrying its signature line or video. */
function insertImage(view: EditorView, image: InlineImage): boolean {
	const node = schema.nodes.image!.create({
		relId: '',
		partName: image.partName,
		contentType: image.contentType,
		widthPx: image.widthPx,
		heightPx: image.heightPx,
		altText: image.altText ?? null,
		signatureLine: image.signatureLine ? JSON.stringify(image.signatureLine) : null,
		webVideo: image.webVideo ? JSON.stringify(image.webVideo) : null,
	});
	view.dispatch(view.state.tr.replaceSelectionWith(node, false).scrollIntoView());
	return true;
}

/**
 * The signature line's picture: a cross and the signing line, with the suggested signer and title
 * under it as Word draws them. Needs a 2D canvas; without one the plain placeholder is used.
 */
async function signaturePicture(spec: SignatureLineSpec): Promise<Uint8Array> {
	const width = SIGNATURE_LINE_SIZE_PT.width * 2;
	const height = SIGNATURE_LINE_SIZE_PT.height * 2;
	const canvas = typeof document === 'undefined' ? undefined : document.createElement('canvas');
	const context = canvas?.getContext?.('2d');
	if (!canvas || !context || typeof canvas.toBlob !== 'function')
		return signatureLinePlaceholderPng();
	canvas.width = width;
	canvas.height = height;
	const named = Boolean(spec.suggestedSigner || spec.suggestedSignerTitle);
	const base = Math.round(height * (named ? 0.58 : 0.72));
	context.strokeStyle = '#5a5a5a';
	context.fillStyle = '#5a5a5a';
	context.lineWidth = 3;
	context.beginPath();
	context.moveTo(14, base - 40);
	context.lineTo(42, base - 12);
	context.moveTo(42, base - 40);
	context.lineTo(14, base - 12);
	context.stroke();
	context.lineWidth = 2;
	context.beginPath();
	context.moveTo(8, base);
	context.lineTo(width - 8, base);
	context.stroke();
	context.textBaseline = 'top';
	let y = base + 8;
	for (const [text, font] of [
		[spec.suggestedSigner, 'bold 22px sans-serif'],
		[spec.suggestedSignerTitle, '20px sans-serif'],
	] as const) {
		if (!text) continue;
		context.font = font;
		context.fillText(text, 10, y, width - 20);
		y += 28;
	}
	const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, PNG));
	return blob ? new Uint8Array(await blob.arrayBuffer()) : signatureLinePlaceholderPng();
}

/** Insert > Signature Line: the VML signature line with its picture, at the selection. */
export async function insertSignatureLine(
	view: EditorView,
	spec: SignatureLineSpec,
	stage: StageGenerated,
	altText?: string,
): Promise<boolean> {
	if (!view.editable) return false;
	const partName = stage(await signaturePicture(spec), PNG);
	const done = insertImage(view, signatureLineImage(spec, { partName, contentType: PNG }, altText));
	focusView(view);
	return done;
}

/** Insert > Online Video: the poster picture carrying the embed, at the selection. */
export function insertWebVideo(
	view: EditorView,
	spec: WebVideoSpec,
	stage: StageGenerated,
	altText?: string,
): boolean {
	if (!view.editable) return false;
	const partName = stage(webVideoPosterPng(), PNG);
	const done = insertImage(view, webVideoImage(spec, { partName, contentType: PNG }, altText));
	focusView(view);
	return done;
}
