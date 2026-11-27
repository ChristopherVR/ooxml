// Theme picture fills in a Word document: the pictures `word/theme/theme1.xml` names through its
// own relationships are read onto the format scheme, and written back (media part, relationship,
// content type) when the theme part is rewritten, so a theme imported from a `.thmx` keeps them.
import type JSZip from 'jszip';
import { imageContentType } from '../drawingml/image-content-type';
import {
	MAX_THEME_PICTURE_BYTES,
	MAX_THEME_PICTURES_BYTES,
	THEME_PICTURE_TYPES,
	themeFillMediaBytes,
	themeFillMediaOf,
	themeSchemeMedia,
	withThemeSchemeMedia,
	type ThemeFillMedia,
} from '../drawingml/theme-fill-media';
import { embedPackageImage } from '../opc/package-image';
import { RELATIONSHIP_TYPES } from '../opc/relationship-types';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import { NS, buildXml, parseXml } from '../xml/index';
import type { ThemeCatalog } from './theme-model';

/** The theme with each picture fill given the picture its relationship names, within limits. */
export async function readWordThemeMedia(
	zip: JSZip,
	themePath: string,
	theme: ThemeCatalog,
): Promise<ThemeCatalog> {
	const scheme = theme.formatScheme;
	if (!scheme) return theme;
	const wanted = new Set(
		[...scheme.fillStyles, ...scheme.backgroundFillStyles].flatMap((style) =>
			style.fill.kind === 'picture' && style.fill.relId ? [style.fill.relId] : [],
		),
	);
	if (!wanted.size) return theme;
	const rels = parseRelationships(await zip.file(relationshipsPartFor(themePath))?.async('string'));
	const media = new Map<string, ThemeFillMedia>();
	let total = 0;
	for (const id of wanted) {
		const rel = rels.get(id);
		if (!rel || rel.mode !== 'Internal' || rel.type !== RELATIONSHIP_TYPES.image) continue;
		const target = resolvePartPath(themePath, rel.target);
		const contentType = imageContentType(target);
		const file = zip.file(target);
		if (!file || !(contentType in THEME_PICTURE_TYPES)) continue;
		const bytes = await file.async('uint8array');
		if (bytes.byteLength > MAX_THEME_PICTURE_BYTES) continue;
		total += bytes.byteLength;
		if (total > MAX_THEME_PICTURES_BYTES) break;
		media.set(id, themeFillMediaOf(bytes, contentType));
	}
	return media.size
		? { ...theme, formatScheme: withThemeSchemeMedia(scheme, (id) => media.get(id)) }
		: theme;
}

/**
 * Stores the pictures of the theme's picture fills for the theme part at `themePath` and points
 * the part's `r:embed` ids at them. A picture the part already relates to keeps its relationship.
 */
export async function writeWordThemeMedia(
	zip: JSZip,
	themePath: string,
	theme: ThemeCatalog,
): Promise<void> {
	const media = themeSchemeMedia(theme.formatScheme);
	const file = zip.file(themePath);
	if (!media.size || !file) return;
	const remap = new Map<string, string>();
	for (const [id, picture] of media) {
		const bytes = themeFillMediaBytes(picture);
		if (!bytes) continue;
		const extension = THEME_PICTURE_TYPES[picture.contentType]!;
		remap.set(id, await embedPackageImage(zip, themePath, 'word/media/', { bytes, extension }));
	}
	if ([...remap].every(([from, to]) => from === to)) return;
	const doc = parseXml(await file.async('string'));
	for (const blip of Array.from(doc.getElementsByTagNameNS(NS.a, 'blip'))) {
		const id = blip.getAttributeNS(NS.r, 'embed');
		const next = id ? remap.get(id) : undefined;
		if (next && next !== id) blip.setAttributeNS(NS.r, 'r:embed', next);
	}
	zip.file(themePath, buildXml(doc));
}
