import { resolveDrawingOuterShadow } from 'ooxml-core/drawingml';
import {
	drawingSchemeColors,
	pictureFramePath,
	shapeShadow,
	shapeStroke,
	type ImageObject,
	type ThemePalette,
} from 'ooxml-core/xlsx';
import { h } from './dom';

/** A clipped picture surface and its inherited or authored outer frame. */
export function paintPictureFrame(
	node: HTMLElement,
	picture: ImageObject,
	box: { w: number; h: number },
	scale: number,
	theme: ThemePalette | undefined,
): HTMLElement {
	const doc = node.ownerDocument;
	const path = pictureFramePath(picture, box.w, box.h);
	const surface = h(doc, 'div');
	Object.assign(surface.style, { position: 'absolute', inset: '0', overflow: 'hidden' });
	if (path) surface.style.clipPath = `path("${path}")`;
	node.append(surface);
	const outer = shapeShadow(picture, theme);
	if (outer) {
		const scheme = drawingSchemeColors(theme);
		const shadow = resolveDrawingOuterShadow(outer, { scheme: (name) => scheme[name] });
		if (shadow) {
			const color = `color-mix(in srgb, ${shadow.color} ${shadow.opacity * 100}%, transparent)`;
			const offset = `${shadow.dx * scale}px ${shadow.dy * scale}px`;
			if (path)
				node.style.filter = `drop-shadow(${offset} ${(shadow.blur * scale) / 2}px ${color})`;
			else node.style.boxShadow = `${offset} ${shadow.blur * scale}px ${color}`;
		}
	}
	const { stroke, width, dash } = shapeStroke(picture, scale, theme);
	if (stroke === 'none') return surface;
	if (path) {
		const ns = 'http://www.w3.org/2000/svg';
		const border = doc.createElementNS(ns, 'svg');
		border.setAttribute('width', String(box.w));
		border.setAttribute('height', String(box.h));
		border.setAttribute('aria-hidden', 'true');
		Object.assign(border.style, {
			position: 'absolute',
			inset: '0',
			pointerEvents: 'none',
			overflow: 'visible',
		});
		const outline = doc.createElementNS(ns, 'path');
		outline.setAttribute('d', path);
		outline.setAttribute('fill', 'none');
		outline.setAttribute('stroke', stroke);
		outline.setAttribute('stroke-width', String(width));
		if (dash) outline.setAttribute('stroke-dasharray', dash);
		border.append(outline);
		node.append(border);
	} else {
		const border = h(doc, 'span', '', { 'aria-hidden': 'true' });
		Object.assign(border.style, {
			position: 'absolute',
			inset: '0',
			pointerEvents: 'none',
			boxShadow: `inset 0 0 0 ${width}px ${stroke}`,
		});
		node.append(border);
	}
	return surface;
}
