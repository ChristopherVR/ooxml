// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Insert > Online Video: the picture Word writes for a video that plays from a web page. The
// picture is a poster, and `wp15:webVideoPr` on its `wp:docPr` carries the iframe that plays. No
// request is made here: the poster is drawn locally, and the iframe is only text in the file. A
// host chooses which video sites may be embedded.
import { encodePng } from '../drawingml/raster-encode';
import type { InlineImage } from './inline-model';
import { makeNS, type XmlDocument, type XmlElement } from './xml';

export const WP15_NS = 'http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
/** The `a:ext` URI Word gives `wp15:webVideoPr`. */
export const WEB_VIDEO_EXT_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';

export interface WebVideoSpec {
	/** The `https:` page the iframe plays. */
	url: string;
	/** The iframe markup Word keeps in `embeddedHtml`. */
	embeddedHtml: string;
	/** The player's size in pixels. */
	width: number;
	height: number;
}

export type WebVideoResult =
	| { ok: true; spec: WebVideoSpec }
	| { ok: false; reason: 'invalid' | 'host' };

/** The size Word gives a new online video's picture, in CSS pixels. */
export const WEB_VIDEO_SIZE_PX = { width: 480, height: 270 } as const;

/** The `src` of an `<iframe>` embed code, or the text itself when it is a plain address. */
function embedAddress(input: string): string | undefined {
	const text = input.trim();
	const tag = /^<iframe\b[^>]*>/i.exec(text);
	if (!tag) return /^\S+$/.test(text) ? text : undefined;
	const src = /\ssrc\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(tag[0]);
	return src?.[1] ?? src?.[2];
}

/** Whether `hostname` is `entry` or a subdomain of it; entries are bare host names. */
const hostAllowed = (hostname: string, allowList: readonly string[]): boolean =>
	allowList.some((entry) => {
		const host = entry.trim().toLowerCase().replace(/^\*\./, '');
		return host !== '' && (hostname === host || hostname.endsWith(`.${host}`));
	});

/**
 * An embed code or embed address, checked against the host's allow-list. Only `https:` addresses of
 * an allowed site are accepted; nothing is fetched.
 */
export function resolveWebVideo(input: string, allowList: readonly string[]): WebVideoResult {
	const address = embedAddress(input);
	let url: URL;
	try {
		url = new URL(address ?? '');
	} catch {
		return { ok: false, reason: 'invalid' };
	}
	if (url.protocol !== 'https:' || url.username || url.password)
		return { ok: false, reason: 'invalid' };
	if (!hostAllowed(url.hostname.toLowerCase(), allowList)) return { ok: false, reason: 'host' };
	const src = url.href.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
	return {
		ok: true,
		spec: {
			url: url.href,
			width: 560,
			height: 315,
			embeddedHtml: `<iframe width="560" height="315" src="${src}" frameborder="0" allowfullscreen=""></iframe>`,
		},
	};
}

/** `wp:docPr` content for an online video: the `a:extLst` that holds `wp15:webVideoPr`. */
export function applyWebVideo(doc: XmlDocument, docPr: XmlElement, spec: WebVideoSpec): void {
	const list = makeNS(doc, A_NS, 'a:extLst');
	const ext = makeNS(doc, A_NS, 'a:ext');
	ext.setAttribute('uri', WEB_VIDEO_EXT_URI);
	const video = makeNS(doc, WP15_NS, 'wp15:webVideoPr');
	video.setAttribute('embeddedHtml', spec.embeddedHtml);
	video.setAttribute('h', String(spec.height));
	video.setAttribute('w', String(spec.width));
	ext.appendChild(video);
	list.appendChild(ext);
	docPr.appendChild(list);
}

/** The online video a `wp:docPr` describes, if it carries `wp15:webVideoPr`. */
export function parseWebVideo(docPr: XmlElement | undefined): WebVideoSpec | undefined {
	const video = docPr && Array.from(docPr.getElementsByTagNameNS(WP15_NS, 'webVideoPr'))[0];
	const html = video?.getAttribute('embeddedHtml');
	if (!video || !html) return undefined;
	const src = /\ssrc\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(html);
	const width = Number(video.getAttribute('w'));
	const height = Number(video.getAttribute('h'));
	return {
		url: (src?.[1] ?? src?.[2] ?? '').replace(/&quot;/g, '"').replace(/&amp;/g, '&'),
		embeddedHtml: html,
		width: Number.isFinite(width) ? width : 0,
		height: Number.isFinite(height) ? height : 0,
	};
}

/** A new online video's picture reference and size, for the run model. */
export function webVideoImage(
	spec: WebVideoSpec,
	picture: Pick<InlineImage, 'partName' | 'contentType'>,
	altText = 'Online video',
): InlineImage {
	return {
		relId: '',
		partName: picture.partName,
		contentType: picture.contentType,
		widthPx: WEB_VIDEO_SIZE_PX.width,
		heightPx: WEB_VIDEO_SIZE_PX.height,
		altText,
		webVideo: spec,
	};
}

/** The poster of a new online video: a dark frame with a play triangle, drawn locally. */
export function webVideoPosterPng(): Uint8Array {
	const width = 320;
	const height = 180;
	const rgb = new Uint8Array(width * height * 3);
	for (let at = 0; at < width * height; at++) rgb.fill(38, at * 3, at * 3 + 3);
	const cx = width / 2;
	const cy = height / 2;
	// A right-pointing triangle: at row y the fill runs from the left edge to a point that narrows.
	for (let y = -34; y <= 34; y++)
		for (let x = -28; x <= 28 - Math.abs(y); x++) {
			const at = (Math.round(cy + y) * width + Math.round(cx + x)) * 3;
			rgb.fill(235, at, at + 3);
		}
	return encodePng(width, height, rgb);
}
