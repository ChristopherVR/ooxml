/**
 * Picture Format > Picture Layout: selected pictures become a SmartArt graphic in one of
 * PowerPoint's built-in picture layouts, each picture in its own node's picture placeholder.
 *
 * The diagram is new and fabricated on save, so its pictures are kept as a picture-layout source
 * (`DiagramLayoutSource.pictures`), the same association a loaded picture diagram keeps: the
 * placeholder-aware engine draws them now, and the save stores them in the data part, the
 * drawing cache and the parts' relationships (`smartart-fabrication-pictures`). Captions start
 * empty, as PowerPoint leaves them ("[Text]"). The layouts come from the built-in definitions,
 * so this module belongs to the lazy `ooxml-core/pptx/smartart-layouts` entry.
 *
 * @module utils/smartart-builtin-layouts/picture-layout
 */
import {
	buildFabricatedColorsXml,
	buildFabricatedQuickStyleXml,
} from '../../../../diagram/write/fabricated-styles';
import type { DiagramPictureSource } from '../../../../diagram/picture-layout';
import type {
	PptxImageLikeElement,
	PptxSmartArtData,
	PptxSmartArtNode,
	SmartArtPptxElement,
} from '../../types';
import { regenerateSmartArtDrawingShapes } from '../smartart-relayout';
import {
	applyBuiltinSmartArtLayout,
	findBuiltinSmartArtLayout,
	listBuiltinSmartArtLayouts,
	type BuiltinSmartArtLayoutEntry,
} from './index';

/** Package-part key of a picture that has no media part yet (saved from its data URL). */
export const PENDING_PICTURE_LAYOUT_PART = 'ppt/media/ooxml-picture-layout-';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The built-in layouts with picture placeholders, in gallery order. */
export function listPictureLayouts(): BuiltinSmartArtLayoutEntry[] {
	return listBuiltinSmartArtLayouts().filter(
		(entry) => entry.category === 'picture' || /\b(?:Picture|Pictures)\b/u.test(entry.title),
	);
}

export interface PictureLayoutOptions {
	/** The new SmartArt element's id. */
	id: string;
	/** A displayable URL for each picture id whose `imageData` is missing. */
	imageUrls?: ReadonlyMap<string, string> | undefined;
	/** The theme's colour map (scheme key to hex) the diagram's colours resolve against. */
	scheme?: Record<string, string> | undefined;
	minorFont?: string | undefined;
}

/** Why pictures cannot be converted, when they cannot. */
export type PictureLayoutRefusal = 'noPictures' | 'unknownLayout' | 'pictureUnavailable';

function dataUrlExtension(url: string): string | undefined {
	const type = /^data:image\/([a-z0-9.+-]+)[;,]/iu.exec(url)?.[1]?.toLowerCase();
	if (!type) return undefined;
	return type === 'jpeg' ? 'jpg' : type === 'svg+xml' ? 'svg' : type;
}

function pictureFillXml(relId: string): string {
	return (
		`<a:blipFill xmlns:a="${A_NS}" xmlns:r="${R_NS}" rotWithShape="1">` +
		`<a:blip r:embed="${relId}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
	);
}

/** Pictures in reading order: top to bottom, then left to right. */
function readingOrder(pictures: readonly PptxImageLikeElement[]): PptxImageLikeElement[] {
	return [...pictures].sort((a, b) => a.y - b.y || a.x - b.x);
}

/** Convert `pictures` into a SmartArt element using the built-in layout `layoutId`. */
export function pictureLayoutSmartArt(
	pictures: readonly PptxImageLikeElement[],
	layoutId: string,
	options: PictureLayoutOptions,
): SmartArtPptxElement | PictureLayoutRefusal {
	if (pictures.length === 0) return 'noPictures';
	if (!findBuiltinSmartArtLayout(layoutId)) return 'unknownLayout';
	const ordered = readingOrder(pictures);
	const nodes: PptxSmartArtNode[] = [];
	const sources: DiagramPictureSource[] = [];
	const pictureParts: Record<string, string> = {};
	const pictureUrls: Record<string, string> = {};
	for (const [index, picture] of ordered.entries()) {
		const url = picture.imageData ?? options.imageUrls?.get(picture.id);
		const extension = url ? dataUrlExtension(url) : undefined;
		// An edited or inserted picture is saved from its bytes; an unchanged one keeps its part.
		const part = extension
			? `${PENDING_PICTURE_LAYOUT_PART}${index + 1}.${extension}`
			: picture.imagePath && !/^(?:[a-z][a-z\d+.-]*:|\/\/)/iu.test(picture.imagePath)
				? picture.imagePath
				: undefined;
		if (!part || !url) return 'pictureUnavailable';
		const nodeId = `${options.id}-node-${index + 1}`;
		const relId = `ooxmlDiagramPicture${index + 1}`;
		nodes.push({ id: nodeId, text: '' });
		sources.push({
			nodeId,
			fill: { kind: 'picture', relId, sourceXml: pictureFillXml(relId) },
		});
		pictureParts[relId] = part;
		pictureUrls[part] = url;
	}
	const left = Math.min(...ordered.map((picture) => picture.x));
	const top = Math.min(...ordered.map((picture) => picture.y));
	const width = Math.max(1, Math.max(...ordered.map((p) => p.x + p.width)) - left);
	const height = Math.max(1, Math.max(...ordered.map((p) => p.y + p.height)) - top);
	const base: PptxSmartArtData = { colorScheme: 'colorful1', style: 'flat', nodes };
	const laidOut = applyBuiltinSmartArtLayout(base, layoutId);
	const data: PptxSmartArtData = {
		...laidOut,
		pictureLayoutSource: {
			...(laidOut.layoutDefinition?.rawXmlText && {
				layoutXml: laidOut.layoutDefinition.rawXmlText,
			}),
			colorsXml: buildFabricatedColorsXml(base.colorScheme),
			quickStyleXml: buildFabricatedQuickStyleXml(base.style),
			pictures: sources,
			pictureParts,
		},
		pictureUrls,
		...(options.scheme && { pictureScheme: options.scheme }),
		...(options.minorFont && { themeMinorFont: options.minorFont }),
	};
	data.drawingShapes = regenerateSmartArtDrawingShapes(data, width, height);
	return {
		type: 'smartArt',
		id: options.id,
		name: 'SmartArt',
		x: left,
		y: top,
		width,
		height,
		smartArtData: data,
	};
}
