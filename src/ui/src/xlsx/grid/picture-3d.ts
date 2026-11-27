// Picture Style 3-D in the worksheet grid (bevels, 3-D rotation and perspective presets), drawn
// with the PowerPoint renderer's shared 3-D pieces rather than a copy: the camera of the picture's
// `a:scene3d` (a COM-measured preset homography, or an explicit rotation) through
// `getCameraTransform`, and its `a:sp3d` bevels lit by `getBevelLightingFilterMarkup` (an SVG
// lighting filter from the light rig and material). Loaded only when a sheet shows such a picture.
//
// An approximation, like PowerPoint's DOM path: the picture is projected as one flat face, so there
// are no extrusion sides, contour or hidden surfaces, the perspective presets' off-axis origin is
// not applied, and a picture that is also rotated in 2-D is turned after the camera, not before.
import type { ImageObject } from 'ooxml-core/xlsx';
import { getBevelLightingFilterMarkup } from '../../pptx/render/visual-3d-bevel-lighting';
import { getCameraTransform } from '../../pptx/render/visual-3d-camera';

let serial = 0;

const scaled = (value: number | undefined, scale: number) =>
	value === undefined ? undefined : value * scale;

/** The CSS camera transform of `picture` at `box`, or undefined when the camera is face-on. */
export function pictureCameraCss(
	picture: Pick<ImageObject, 'scene3d'>,
	box: { w: number; h: number },
): { transform: string; origin?: string } | undefined {
	if (!picture.scene3d) return undefined;
	const camera = getCameraTransform(picture.scene3d, { width: box.w, height: box.h });
	if (camera.matrix3d)
		return { transform: camera.matrix3d, origin: camera.transformOrigin ?? '0 0' };
	if (camera.cameraFlatFace) return undefined;
	const turns = [
		camera.rotateX ? `rotateX(${camera.rotateX}deg)` : '',
		camera.rotateY ? `rotateY(${camera.rotateY}deg)` : '',
		camera.rotateZ ? `rotateZ(${camera.rotateZ}deg)` : '',
	].filter(Boolean);
	if (!turns.length) return undefined;
	return {
		transform: [camera.perspective ? `perspective(${camera.perspective})` : '', ...turns]
			.filter(Boolean)
			.join(' '),
	};
}

/**
 * Wraps what `node` shows of a picture in one face that carries the bevel filter and the camera
 * transform. Returns what it applied, for the tests and the unapplied-effects note.
 */
export function paintPicture3d(
	node: HTMLElement,
	picture: ImageObject,
	box: { w: number; h: number },
	scale: number,
): { camera: boolean; bevel: boolean } {
	const doc = node.ownerDocument;
	const face = doc.createElement('div');
	face.className = 'xg-pic3d';
	Object.assign(face.style, { position: 'absolute', inset: '0' });
	face.append(...node.childNodes);
	node.append(face);
	const shape = picture.shape3d;
	const filter =
		shape &&
		getBevelLightingFilterMarkup(
			`xlsx-picture-${++serial}`,
			{
				...shape,
				bevelTopWidth: scaled(shape.bevelTopWidth, scale),
				bevelTopHeight: scaled(shape.bevelTopHeight, scale),
				bevelBottomWidth: scaled(shape.bevelBottomWidth, scale),
				bevelBottomHeight: scaled(shape.bevelBottomHeight, scale),
			},
			picture.scene3d,
		);
	if (filter) {
		const ns = 'http://www.w3.org/2000/svg';
		const svg = doc.createElementNS(ns, 'svg');
		svg.setAttribute('width', '0');
		svg.setAttribute('height', '0');
		svg.setAttribute('aria-hidden', 'true');
		Object.assign(svg.style, { position: 'absolute' });
		const defs = doc.createElementNS(ns, 'defs');
		// The markup is built from numbers and an escaped id only, as PowerPoint's filter defs are.
		defs.innerHTML = filter.filterMarkup;
		svg.append(defs);
		node.append(svg);
		face.style.filter = filter.cssReference;
	}
	const camera = pictureCameraCss(picture, box);
	if (camera) {
		face.style.transform = camera.transform;
		if (camera.origin) face.style.transformOrigin = camera.origin;
	}
	return { camera: !!camera, bevel: !!filter };
}
