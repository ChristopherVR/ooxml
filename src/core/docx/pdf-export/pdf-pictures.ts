/**
 * Pictures for the Word PDF export: JPEG bytes go into the file as they are (no re-encoding,
 * original resolution); PNG is decoded to its samples and alpha. Other formats go through a
 * decoder the host supplies, or are left out and reported. Charts, drawing shapes and text boxes
 * go to `drawings` (`pdf-drawings.ts`). DOM-free.
 */
import { decodePng, jpegComponents, inspectRasterImage } from '../../drawingml/index';
import type { LayoutObject } from '../layout/index';
import { drawBlurredSilhouette } from '../../pdf/write/pdf-blur-mask';
import { parsePath } from '../../pdf/write/pdf-path';
import type { PdfImageResource, PdfVectorPage } from '../../pdf/write/pdf-vector';
import type { DocxPdfDrawings } from './pdf-drawings';
import type { DocxPdfReportBuilder } from './pdf-report';
import { cssRgba } from './pdf-text';

/** RGBA pixels a host decoder returns for a picture format the export cannot read itself. */
export interface DecodedPicture {
	width: number;
	height: number;
	/** 8-bit RGBA, `width * height * 4` bytes. */
	rgba: Uint8Array | Uint8ClampedArray;
}

export type DocxPdfMedia =
	| ReadonlyMap<string, Uint8Array>
	| ((partName: string) => Uint8Array | undefined);

export type DocxPdfPictureDecoder = (
	bytes: Uint8Array,
	contentType: string,
) => Promise<DecodedPicture | undefined> | DecodedPicture | undefined;

/** A picture to draw: what the layout's inline objects and floats have in common. */
export type DocxPdfPicture = Pick<
	LayoutObject,
	| 'partName'
	| 'contentType'
	| 'widthPx'
	| 'heightPx'
	| 'pictureFill'
	| 'shadow'
	| 'chart'
	| 'textBox'
	| 'textBoxLayout'
>;

function fromRgba(picture: DecodedPicture): PdfImageResource | undefined {
	const { width, height, rgba } = picture;
	if (!(width > 0 && height > 0) || rgba.length < width * height * 4) return undefined;
	const rgb = new Uint8Array(width * height * 3);
	const alpha = new Uint8Array(width * height);
	let translucent = false;
	for (let pixel = 0; pixel < width * height; pixel++) {
		rgb[pixel * 3] = rgba[pixel * 4]!;
		rgb[pixel * 3 + 1] = rgba[pixel * 4 + 1]!;
		rgb[pixel * 3 + 2] = rgba[pixel * 4 + 2]!;
		alpha[pixel] = rgba[pixel * 4 + 3]!;
		if (alpha[pixel] !== 255) translucent = true;
	}
	return { width, height, data: rgb, ...(translucent ? { alpha: { data: alpha } } : {}) };
}

/** Reads, decodes and caches each picture part once per export. */
export class DocxPdfPictures {
	readonly #resources = new Map<string, PdfImageResource | null>();
	/** Draws charts, shapes and text boxes; without it they are reported as left out. */
	drawings: DocxPdfDrawings | undefined;
	constructor(
		readonly media: DocxPdfMedia | undefined,
		readonly decoder: DocxPdfPictureDecoder | undefined,
		readonly report: DocxPdfReportBuilder,
	) {}

	#bytes(partName: string): Uint8Array | undefined {
		if (!this.media) return undefined;
		return typeof this.media === 'function' ? this.media(partName) : this.media.get(partName);
	}

	/** Loads every picture `pictures` name, so painting stays synchronous. */
	async prepare(pictures: Iterable<DocxPdfPicture>): Promise<void> {
		for (const picture of pictures) {
			if (picture.chart) {
				await this.drawings?.prepare(picture);
				continue;
			}
			if (!picture.partName || this.#resources.has(picture.partName)) continue;
			const bytes = this.#bytes(picture.partName);
			this.#resources.set(
				picture.partName,
				bytes ? await this.decode(bytes, picture.contentType) : null,
			);
		}
	}

	/** A loaded picture part, after `prepare`. */
	resource(partName: string): PdfImageResource | undefined {
		return this.#resources.get(partName) ?? undefined;
	}

	/** Reads picture bytes as an image resource: JPEG as it is, PNG decoded, others by the host. */
	async decode(bytes: Uint8Array, contentType: string): Promise<PdfImageResource | null> {
		let kind: string | undefined;
		try {
			kind = inspectRasterImage(bytes).mimeType;
		} catch {
			kind = undefined;
		}
		if (kind === 'image/jpeg') {
			const components = jpegComponents(bytes);
			const size = inspectRasterImage(bytes);
			if (components === 1 || components === 3)
				return {
					width: size.pixelWidth,
					height: size.pixelHeight,
					data: bytes,
					filter: 'DCTDecode',
					...(components === 1 ? { colorSpace: 'DeviceGray' as const } : {}),
				};
		}
		if (kind === 'image/png') {
			const decoded = decodePng(bytes);
			if (decoded)
				return {
					width: decoded.width,
					height: decoded.height,
					data: decoded.rgb,
					...(decoded.alpha ? { alpha: { data: decoded.alpha } } : {}),
				};
		}
		const decoded = await this.decoder?.(bytes, contentType);
		return (decoded && fromRgba(decoded)) ?? null;
	}

	/** Draws a picture into the box at (x, y), cropped as the document crops it. */
	draw(page: PdfVectorPage, picture: DocxPdfPicture, x: number, y: number): void {
		const { widthPx: width, heightPx: height } = picture;
		// A line or connector has a zero extent across it and is still stroked.
		const line = !!picture.chart?.shape && width >= 0 && height >= 0 && width + height > 0;
		if (!line && !(width > 0 && height > 0)) return;
		if (picture.chart) {
			if (this.drawings) this.drawings.draw(page, picture, x, y);
			else
				this.report.omitted(
					picture.chart.shape ? 'shape' : 'chart',
					'Charts and drawing shapes are not exported to PDF without a drawing painter.',
				);
			return;
		}
		const resource = this.#resources.get(picture.partName);
		if (!resource) {
			this.report.omitted(
				'picture',
				`Picture ${picture.partName} (${picture.contentType}) could not be read.`,
			);
			return;
		}
		if (picture.shadow) {
			const color = cssRgba(picture.shadow.color) ?? { r: 0, g: 0, b: 0, a: 1 };
			const segments = parsePath(`M 0 0 L ${width} 0 L ${width} ${height} L 0 ${height} Z`);
			if (segments)
				drawBlurredSilhouette(
					page,
					[
						{
							segments,
							fill: 'nonzero',
							matrix: [1, 0, 0, 1, x + picture.shadow.dx, y + picture.shadow.dy],
						},
					],
					picture.shadow.blur / 2,
					color,
					picture.shadow.opacity,
				);
		}
		const crop = picture.pictureFill?.crop ?? {};
		const left = crop.cropLeft ?? 0;
		const top = crop.cropTop ?? 0;
		const right = crop.cropRight ?? 0;
		const bottom = crop.cropBottom ?? 0;
		const fullWidth = width / Math.max(0.01, 1 - left - right);
		const fullHeight = height / Math.max(0.01, 1 - top - bottom);
		page.save();
		page.rect(x, y, width, height);
		page.op('W n');
		// The image fills the unit square with y up; the page space is y down.
		page.concat([
			fullWidth,
			0,
			0,
			-fullHeight,
			x - left * fullWidth,
			y - top * fullHeight + fullHeight,
		]);
		page.drawImage(resource);
		page.restore();
	}
}
