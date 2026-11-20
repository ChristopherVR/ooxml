// A picture node's bytes on screen: its SVG original or raster part as the `<img>` source, and the
// SVG drawing that crops or recolours it when its fill needs one. The editing surface's node view
// and the read-only story previews both paint pictures here, from the same media cache.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { pictureNeedsSvg, pictureSvg } from '../picture-view';

/** Resolves a package part to a displayable URL (the editor's `ImageMediaCache.urlFor`). */
export type PictureSource = (partName: string, contentType: string) => string | undefined;

/** The editor's media cache as a picture view uses it: URLs, and a call when they are replaced. */
export interface PictureMedia {
	urlFor(partName: string, contentType: string): string | undefined;
	/** Called after the cache revokes its URLs because bytes arrived; returns the unsubscribe. */
	onChange(listener: () => void): () => void;
}

/**
 * Sets `content`'s source from the node's parts and, when its fill needs SVG, draws that SVG after
 * it (replacing `painted`, the one drawn before) and hides the `<img>`. Returns the SVG drawn.
 */
export function paintPicture(
	content: HTMLElement,
	node: ProseMirrorNode,
	urlFor: PictureSource,
	painted?: SVGSVGElement,
): SVGSVGElement | undefined {
	const src =
		(node.attrs.svgPartName && urlFor(String(node.attrs.svgPartName), 'image/svg+xml')) ||
		urlFor(String(node.attrs.partName), String(node.attrs.contentType));
	if (src) content.setAttribute('src', src);
	content.classList.toggle('dve-image-missing', !src);
	const fill =
		typeof node.attrs.pictureFill === 'string' ? JSON.parse(node.attrs.pictureFill) : undefined;
	if (!src || !pictureNeedsSvg(fill)) return painted;
	const next = pictureSvg(
		content.ownerDocument,
		fill,
		src,
		Number(node.attrs.widthPx),
		Number(node.attrs.heightPx),
	);
	if (painted) painted.replaceWith(next);
	else content.after(next);
	content.style.display = 'none';
	return next;
}
