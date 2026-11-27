// Theme picture fills (`a:blipFill` in a theme's fill or background fill styles) refer to pictures
// through the theme part's own relationships. The format scheme carries each picture with its
// style, so a theme moved between packages (a `.thmx`, a document) keeps its picture fills: the
// writer that stores the theme part recreates the relationship and the media part.
import { decodeBase64, encodeBase64 } from '../digest/base64';
import { NS, elements, named, type XmlElement } from '../xml/index';
import type { ThemeFillStyle, ThemeFormatScheme } from './format-scheme';

/** A picture a theme fill style uses: its content type and Base64 bytes (JSON-safe). */
export interface ThemeFillMedia {
	contentType: string;
	data: string;
}

/** The picture types a theme fill may carry, with the extension its media part takes. */
export const THEME_PICTURE_TYPES: Readonly<Record<string, string>> = {
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/tiff': 'tiff',
};

/** The largest picture one theme fill may carry, and all of a theme's pictures together. */
export const MAX_THEME_PICTURE_BYTES = 4 * 1024 * 1024;
export const MAX_THEME_PICTURES_BYTES = 8 * 1024 * 1024;

/** The `r:embed` id of an `a:blipFill`, or undefined for any other fill or a linked picture. */
export function themePictureEmbedId(fill: XmlElement): string | undefined {
	if (!named(fill, 'blipFill', NS.a, false)) return undefined;
	const blip = elements(fill).find((node) => named(node, 'blip', NS.a, false));
	if (!blip || blip.hasAttributeNS(NS.r, 'link')) return undefined;
	return blip.getAttributeNS(NS.r, 'embed') || undefined;
}

export const themeFillMediaOf = (bytes: Uint8Array, contentType: string): ThemeFillMedia => ({
	contentType,
	data: encodeBase64(bytes),
});

/** The bytes of a theme picture, or undefined when its type or data is not one a theme may hold. */
export function themeFillMediaBytes(media: ThemeFillMedia): Uint8Array | undefined {
	if (!(media.contentType in THEME_PICTURE_TYPES)) return undefined;
	const bytes = decodeBase64(media.data);
	return bytes && bytes.byteLength <= MAX_THEME_PICTURE_BYTES ? bytes : undefined;
}

const pictureStyles = (scheme: ThemeFormatScheme): ThemeFillStyle[] =>
	[...scheme.fillStyles, ...scheme.backgroundFillStyles].filter(
		(style) => style.fill.kind === 'picture' && style.fill.relId,
	);

/** Every picture the scheme's fill styles carry, by the relationship id their XML names. */
export function themeSchemeMedia(
	scheme: ThemeFormatScheme | undefined,
): Map<string, ThemeFillMedia> {
	const media = new Map<string, ThemeFillMedia>();
	if (!scheme) return media;
	for (const style of pictureStyles(scheme)) {
		const id = style.fill.kind === 'picture' ? style.fill.relId : undefined;
		if (id && style.media) media.set(id, style.media);
	}
	return media;
}

/** The scheme with each picture fill style given the picture `resolve` finds for its id. */
export function withThemeSchemeMedia(
	scheme: ThemeFormatScheme,
	resolve: (relId: string) => ThemeFillMedia | undefined,
): ThemeFormatScheme {
	const attach = (style: ThemeFillStyle): ThemeFillStyle => {
		if (style.fill.kind !== 'picture' || !style.fill.relId || style.media) return style;
		const media = resolve(style.fill.relId);
		return media ? { ...style, media } : style;
	};
	return {
		...scheme,
		fillStyles: scheme.fillStyles.map(attach),
		backgroundFillStyles: scheme.backgroundFillStyles.map(attach),
	};
}
