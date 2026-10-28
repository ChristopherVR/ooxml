import { expect, it } from 'vitest';
import { NS, parseXml, buildXml } from '../xml/index';
import { applyPictureCrop, parsePictureCrop } from './picture-crop';
import { imageRectFromCrop, cropUpdateFromRects } from '../geometry/picture-crop';
it('preserves signed and tiny crop percentages and opaque attributes when patching', () => {
	const fill = parseXml(
		`<a:blipFill xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:blip/><a:srcRect l="-12.5%" r="2" v:opaque="keep"/><a:stretch/></a:blipFill>`,
	).documentElement;
	expect(parsePictureCrop(fill)).toEqual({ cropLeft: -0.125, cropRight: 0.00002 });
	applyPictureCrop(fill, { cropLeft: -0.2, cropRight: 0.00002 });
	expect(buildXml(fill)).toContain('v:opaque="keep"');
	expect(parsePictureCrop(fill)).toEqual({ cropLeft: -0.2, cropRight: 0.00002 });
	const frame = { x: 10, y: 20, width: 200, height: 100 };
	const insets = { cropLeft: 0.25, cropTop: 0, cropRight: -0.25, cropBottom: 0 };
	expect(cropUpdateFromRects(frame, imageRectFromCrop(frame, insets))).toEqual({
		...frame,
		...insets,
	});
});
