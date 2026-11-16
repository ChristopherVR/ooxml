/**
 * Relationship allocation for chart picture fills (`a:blipFill` on a chart element), shared by
 * every host that saves a chart part. A Format pane hands a picture over as a `data:` URL; the host
 * calls {@link embedChartPictures} with its chart part and media folder before it writes the
 * fills, and puts each returned id in `a:blip/@r:embed` (`chartFillFromAppearance` with
 * `picture.relId`). The same URL twice, or a picture the chart already references, shares one
 * media part and relationship.
 *
 * Which hosts offer picture fills is a separate admission decision: the shared chart painter does
 * not draw picture fills yet, so the `ChartSpace` Format pane source keeps the choice disabled and
 * only PowerPoint, which paints them, offers it.
 *
 * @module chart/picture-parts
 */
import type JSZip from 'jszip';
import { decodeImageDataUrl, embedPackageImage, type PackageImage } from '../opc/package-image';

/** The media folder of each host's package, where chart pictures are stored. */
export const CHART_MEDIA_FOLDERS = {
	pptx: 'ppt/media/',
	docx: 'word/media/',
	xlsx: 'xl/media/',
	visio: 'visio/media/',
} as const;

/**
 * Stores each picture for `chartPartName` and returns the relationship id of each source. A
 * source is a `data:image/...` URL or decoded bytes keyed by the caller; a URL that does not
 * decode is left out of the result, so the caller keeps that fill unchanged.
 */
export async function embedChartPictures(
	zip: JSZip,
	chartPartName: string,
	mediaFolder: string,
	sources: Iterable<string>,
	decode: (source: string) => PackageImage | undefined | null = decodeImageDataUrl,
): Promise<Map<string, string>> {
	const ids = new Map<string, string>();
	for (const source of sources) {
		if (ids.has(source)) continue;
		const image = decode(source);
		if (image) ids.set(source, await embedPackageImage(zip, chartPartName, mediaFolder, image));
	}
	return ids;
}
