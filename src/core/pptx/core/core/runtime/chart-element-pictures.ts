/**
 * Picture fills of chart element formats (`a:blipFill` on the chart area,
 * plot area, legend, title, a series, ...), both directions:
 *
 * - load: a record read off the part names its image by relationship id;
 *   {@link resolveChartElementPictures} turns that into a `data:` URL the
 *   renderer paints;
 * - save: a picture chosen in the Format pane has only its `data:` URL;
 *   {@link embedChartElementPictures} writes it as a new media part, adds the
 *   chart part relationship and returns the formats with the new ids, which
 *   the element format writer puts in `a:blip/@r:embed`.
 *
 * @module runtime/chart-element-pictures
 */
import type JSZip from 'jszip';

import { CHART_MEDIA_FOLDERS, embedChartPictures } from '../../../../chart/picture-parts';
import type { PptxChartElementFormats, XmlObject } from '../../types';

type ChartRel = { id: string; type: string; target: string };

/** Fill in `src` for every picture record that only names a relationship. Mutates `formats`. */
export async function resolveChartElementPictures(
	formats: PptxChartElementFormats | undefined,
	chartPartPath: string,
	readChartRels: (path: string) => Promise<ChartRel[]>,
	resolveImagePath: (basePath: string, target: string) => string,
	getImageData: (imagePath: string) => Promise<string | undefined>,
): Promise<void> {
	let rels: ChartRel[] | undefined;
	for (const record of Object.values(formats ?? {})) {
		const picture = record.fill?.kind === 'picture' ? record.fill.picture : undefined;
		if (!picture?.relId || picture.src) continue;
		rels ??= await readChartRels(chartPartPath);
		const rel = rels.find((candidate) => candidate.id === picture.relId);
		if (!rel) continue;
		const src = await getImageData(resolveImagePath(chartPartPath, rel.target));
		if (src) picture.src = src;
	}
}

interface EmbedHost {
	zip: JSZip;
	/** Kept for callers; relationships and content types go through the shared package writer. */
	parse?: (xml: string) => XmlObject;
	build?: (tree: XmlObject) => string;
	parseDataUrl: (dataUrl: string) => { bytes: Uint8Array; extension: string } | null;
}

/**
 * Embed every picture the formats carry without a relationship id through the shared chart
 * picture allocation (`chart/picture-parts.ts`): write the media part and the relationship, and
 * return the formats with `relId` set (the input is left alone). Returns the input when there is
 * nothing to embed.
 */
export async function embedChartElementPictures(
	formats: PptxChartElementFormats | undefined,
	chartPartPath: string,
	host: EmbedHost,
): Promise<PptxChartElementFormats | undefined> {
	const pending = Object.entries(formats ?? {}).filter(
		([, record]) =>
			record.fill?.kind === 'picture' && record.fill.picture?.src && !record.fill.picture.relId,
	);
	if (!formats || pending.length === 0) return formats;
	const ids = await embedChartPictures(
		host.zip,
		chartPartPath,
		CHART_MEDIA_FOLDERS.pptx,
		pending.map(([, record]) => record.fill!.picture!.src!),
		host.parseDataUrl,
	);
	const out: PptxChartElementFormats = { ...formats };
	for (const [key, record] of pending) {
		const picture = record.fill!.picture!;
		const id = ids.get(picture.src!);
		if (id) out[key] = { ...record, fill: { ...record.fill!, picture: { ...picture, relId: id } } };
	}
	return out;
}

/** Declare the relationships namespace on `chartSpace` so `r:embed` resolves. */
export function ensureRelationshipNamespace(chartSpace: XmlObject): void {
	if (!chartSpace['@_xmlns:r']) {
		chartSpace['@_xmlns:r'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	}
}
