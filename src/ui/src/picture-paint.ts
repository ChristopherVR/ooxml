import {
	imageRectFromCrop,
	pictureTileGeometry,
	pictureFillCounterTransform,
} from 'ooxml-core/geometry';
import type { DrawingFill } from 'ooxml-core/drawingml';
import { pictureEffectFilter } from './picture-effects';
import { getCachedNativeImageSize, probeNativeImageSize } from './image-native-size';
export interface PictureMedia {
	url: string;
	width?: number;
	height?: number;
}
export interface PicturePaintContext {
	scheme?: Readonly<Record<string, string>>;
	width: number;
	height: number;
	rotation?: number;
	flipHorizontal?: boolean;
	flipVertical?: boolean;
}
const NS = 'http://www.w3.org/2000/svg';
let serial = 0;
/** Shared SVG picture paint, clipped in source space before tile/mirror placement. */
export function picturePaint(
	doc: Document,
	defs: SVGDefsElement,
	fill: Extract<DrawingFill, { kind: 'picture' }>,
	media: string | PictureMedia,
	context: PicturePaintContext = { width: 1, height: 1 },
): { paint: string; unapplied: string[]; approximated?: string[] } {
	if (fill.unsupportedPicture?.length)
		return { paint: '#d1d5db', unapplied: fill.unsupportedPicture };
	const source = typeof media === 'string' ? { url: media } : media;
	const pattern = doc.createElementNS(NS, 'pattern');
	pattern.id = `office-picture-${++serial}`;
	pattern.setAttribute('patternUnits', 'userSpaceOnUse');
	const unapplied: string[] = [];
	const crop = { cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0, ...fill.crop };
	const spanX = 1 - crop.cropLeft - crop.cropRight,
		spanY = 1 - crop.cropTop - crop.cropBottom;
	const counter =
		fill.layout?.rotWithShape === false
			? pictureFillCounterTransform(
					context.width,
					context.height,
					context.rotation,
					context.flipHorizontal,
					context.flipVertical,
				)
			: undefined;
	if (counter) pattern.setAttribute('patternTransform', counter);
	const effects = pictureEffectFilter(doc, defs, fill.effects ?? [], context.scheme);
	unapplied.push(...effects.unapplied);
	if (effects.approximated?.length)
		pattern.dataset.approximatedPictureEffects = effects.approximated.join(',');
	const appendImage = (parent: SVGElement, x: number, y: number, width: number, height: number) => {
		const image = doc.createElementNS(NS, 'image');
		for (const [key, value] of Object.entries({ x, y, width, height }))
			image.setAttribute(key, String(value));
		image.setAttribute('href', source.url);
		image.setAttribute('preserveAspectRatio', 'none');
		if (effects.filter) image.setAttribute('filter', effects.filter);
		parent.append(image);
	};
	if (spanX <= 0 || spanY <= 0) {
		unapplied.push('invalid-crop');
		pattern.dataset.unapplied = 'invalid-crop';
	} else if (fill.layout?.tile) {
		pattern.setAttribute('patternContentUnits', 'userSpaceOnUse');
		const paint = (size: { width: number; height: number } | undefined) => {
			if (!size) {
				unapplied.push('tile-image-size');
				pattern.dataset.unapplied = 'tile-image-size';
				return;
			}
			const tile = fill.layout!.tile!;
			const cell = pictureTileGeometry(
				context,
				// Office scales the uncropped bitmap to set the tile period; srcRect
				// expands its image inside that fixed cell (native PowerPoint probe).
				size,
				tile,
				fill.layout?.dpi,
			);
			if (!cell) {
				unapplied.push('invalid-tile');
				pattern.dataset.unapplied = 'invalid-tile';
				return;
			}
			pattern.replaceChildren();
			for (const [key, value] of Object.entries({
				x: cell.x,
				y: cell.y,
				width: cell.width * cell.columns,
				height: cell.height * cell.rows,
			}))
				pattern.setAttribute(key, String(value));
			for (let row = 0; row < cell.rows; row++)
				for (let col = 0; col < cell.columns; col++) {
					const clip = doc.createElementNS(NS, 'clipPath');
					clip.id = `office-picture-clip-${++serial}`;
					const rect = doc.createElementNS(NS, 'rect');
					rect.setAttribute('width', String(cell.width));
					rect.setAttribute('height', String(cell.height));
					clip.append(rect);
					defs.append(clip);
					const group = doc.createElementNS(NS, 'g');
					group.setAttribute(
						'transform',
						`translate(${(col + (col % 2)) * cell.width} ${(row + (row % 2)) * cell.height}) scale(${col % 2 ? -1 : 1} ${row % 2 ? -1 : 1})`,
					);
					group.setAttribute('clip-path', `url(#${clip.id})`);
					const image = imageRectFromCrop(
						{ x: 0, y: 0, width: cell.width, height: cell.height },
						crop,
					);
					appendImage(group, image.x, image.y, image.width, image.height);
					pattern.append(group);
				}
			pattern.dataset.ready = 'true';
		};
		const size =
			source.width && source.height
				? { width: source.width, height: source.height }
				: getCachedNativeImageSize(source.url);
		if (size) paint(size);
		else void probeNativeImageSize(source.url).then(paint);
	} else {
		pattern.setAttribute('width', String(context.width));
		pattern.setAttribute('height', String(context.height));
		pattern.setAttribute('patternContentUnits', 'objectBoundingBox');
		const insets = fill.layout?.fillRect ?? { l: 0, t: 0, r: 0, b: 0 };
		const frame = {
			x: insets.l,
			y: insets.t,
			width: 1 - insets.l - insets.r,
			height: 1 - insets.t - insets.b,
		};
		if (frame.width <= 0 || frame.height <= 0) {
			unapplied.push('invalid-fillRect');
			pattern.dataset.unapplied = 'invalid-fillRect';
		} else {
			const rect = imageRectFromCrop(frame, crop);
			const clip = doc.createElementNS(NS, 'clipPath');
			clip.id = `office-picture-clip-${++serial}`;
			clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
			const box = doc.createElementNS(NS, 'rect');
			for (const [key, value] of Object.entries(frame)) box.setAttribute(key, String(value));
			clip.append(box);
			defs.append(clip);
			const group = doc.createElementNS(NS, 'g');
			group.setAttribute('clip-path', `url(#${clip.id})`);
			appendImage(group, rect.x, rect.y, rect.width, rect.height);
			pattern.append(group);
		}
	}
	defs.append(pattern);
	return {
		paint: `url(#${pattern.id})`,
		unapplied,
		...(effects.approximated?.length ? { approximated: effects.approximated } : {}),
	};
}
