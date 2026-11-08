/**
 * The Office form of an SVG picture made in code: a raster in `a:blip/@r:embed`
 * and the SVG in an `asvg:svgBlip` extension of that blip (PowerPoint 2016+
 * draws the SVG; older readers draw the raster).
 *
 * @module runtime/picture-svg-blip
 */

import type { XmlObject } from '../../types';

/** `a:ext/@uri` of the SVG blip extension. */
export const SVG_BLIP_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

/** A 1×1 transparent PNG, the raster of an SVG picture given without one. */
export const TRANSPARENT_PNG_DATA_URL =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=';

/** Add the `asvg:svgBlip` extension pointing at `relationshipId` to a picture's blip. */
export function addSvgBlipExtension(picture: XmlObject, relationshipId: string): void {
	const blipFill = picture['p:blipFill'] as XmlObject | undefined;
	const blip = blipFill?.['a:blip'] as XmlObject | undefined;
	if (!blip || typeof blip !== 'object') {
		return;
	}
	const ext: XmlObject = {
		'@_uri': SVG_BLIP_EXT_URI,
		'asvg:svgBlip': { '@_xmlns:asvg': SVG_NS, '@_r:embed': relationshipId },
	};
	const extLst = (blip['a:extLst'] as XmlObject | undefined) ?? {};
	const existing = extLst['a:ext'] as XmlObject | XmlObject[] | undefined;
	const list: XmlObject[] =
		existing === undefined ? [] : Array.isArray(existing) ? existing : [existing];
	extLst['a:ext'] = list.length === 0 ? ext : [...list, ext];
	blip['a:extLst'] = extLst;
}
