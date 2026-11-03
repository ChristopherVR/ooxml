/**
 * The source of a new SVG picture: the SVG as a base64 data URL (what the
 * save writes into its own media part) and the raster the blip embeds.
 *
 * @module sdk/svg-picture-input
 */

import { TRANSPARENT_PNG_DATA_URL } from '../../core/runtime/picture-svg-blip';
import { base64Encode } from '../../utils/modify-verifier-codec';

const SVG_DATA_URL = /^data:image\/svg\+xml(?<params>[^,]*),(?<payload>[\s\S]*)$/iu;

/** The SVG and raster of an SVG picture, or `undefined` for any other source. */
export interface SvgPictureSource {
	/** The SVG as a `data:image/svg+xml;base64,` URL. */
	svgData: string;
	/** The raster `a:blip` embeds, shown where SVG is not read. */
	raster: string;
}

/**
 * Split an `addImage` source into the SVG and its raster when it is an SVG
 * data URL (base64, or percent-encoded / plain text).
 *
 * @throws RangeError for an `svgFallback` that is itself SVG, an
 *   `svgFallback` beside a source that is not SVG, or SVG data that does not
 *   hold an `<svg>` element.
 */
export function svgPictureSource(
	source: string,
	svgFallback: string | undefined,
): SvgPictureSource | undefined {
	if (svgFallback !== undefined && SVG_DATA_URL.test(svgFallback)) {
		throw new RangeError('ImageOptions.svgFallback must be a raster image, not SVG');
	}
	const match = SVG_DATA_URL.exec(source);
	if (!match?.groups) {
		if (svgFallback !== undefined) {
			throw new RangeError('ImageOptions.svgFallback applies only to an SVG data URL source');
		}
		return undefined;
	}
	const params = match.groups['params'] ?? '';
	const payload = match.groups['payload'] ?? '';
	const base64 = /;base64$/iu.test(params);
	const text = base64 ? decodeBase64Text(payload) : decodeURIComponent(payload);
	if (!/<svg[\s>]/iu.test(text)) {
		throw new RangeError('An SVG picture source must hold an <svg> element');
	}
	return {
		svgData: base64
			? `data:image/svg+xml;base64,${payload.replace(/\s+/gu, '')}`
			: `data:image/svg+xml;base64,${base64Encode(new TextEncoder().encode(text))}`,
		raster: svgFallback ?? TRANSPARENT_PNG_DATA_URL,
	};
}

function decodeBase64Text(payload: string): string {
	try {
		const binary = atob(payload.replace(/\s+/gu, ''));
		const bytes = Uint8Array.from(binary, (ch) => ch.charCodeAt(0));
		return new TextDecoder().decode(bytes);
	} catch {
		throw new RangeError('An SVG picture source has base64 data that cannot be read');
	}
}
