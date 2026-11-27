// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { ImageObject } from 'ooxml-core/xlsx';
import { pictureHas3d } from './drawings';
import { paintPicture3d, pictureCameraCss } from './picture-3d';

const picture = (extra: Partial<ImageObject>): ImageObject => ({
	kind: 'image',
	anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	partName: 'xl/media/image1.png',
	contentType: 'image/png',
	...extra,
});

describe('Picture Style 3-D in the grid', () => {
	it('lights a bevel and projects a perspective preset on one face', () => {
		const node = document.createElement('div');
		const surface = document.createElement('div');
		node.append(surface);
		const image = picture({
			scene3d: {
				cameraPreset: 'perspectiveRelaxed',
				lightRigType: 'threePt',
				lightRigDirection: 't',
			},
			shape3d: { bevelTopType: 'circle', bevelTopWidth: 101600, bevelTopHeight: 101600 },
		});
		expect(pictureHas3d(image)).toBe(true);
		expect(paintPicture3d(node, image, { w: 200, h: 120 }, 1)).toEqual({
			camera: true,
			bevel: true,
		});
		const face = node.querySelector<HTMLElement>('.xg-pic3d');
		expect(face?.firstElementChild).toBe(surface);
		expect(face?.style.transform).toMatch(/matrix3d|rotate/);
		// Browsers quote the URL when they serialise it; jsdom may not.
		const id = /^url\("?#(bevel-light-xlsx-picture-\d+)"?\)$/.exec(face?.style.filter ?? '')?.[1];
		expect(id).toBeDefined();
		expect(node.querySelector(`filter[id="${id}"] feSpecularLighting`)).not.toBeNull();
	});

	it('turns an explicit 3-D rotation and leaves face-on cameras flat', () => {
		expect(
			pictureCameraCss(
				picture({ scene3d: { cameraPreset: 'orthographicFront', cameraRotY: 1200000 } }),
				{ w: 100, h: 100 },
			),
		).toBeDefined();
		expect(
			pictureCameraCss(picture({ scene3d: { cameraPreset: 'orthographicFront' } }), {
				w: 100,
				h: 100,
			}),
		).toBeUndefined();
		expect(pictureHas3d(picture({ shape3d: { bevelTopType: 'none' } }))).toBe(false);
		expect(pictureHas3d(picture({}))).toBe(false);
	});
});
