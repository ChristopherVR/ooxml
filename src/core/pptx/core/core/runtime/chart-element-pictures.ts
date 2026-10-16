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

import type { PptxChartElementFormats, XmlObject } from '../../types';

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MIME: Record<string, string> = {
	png: 'image/png',
	jpg: 'image/jpeg',
	jpeg: 'image/jpeg',
	gif: 'image/gif',
	bmp: 'image/bmp',
	svg: 'image/svg+xml',
	webp: 'image/webp',
};

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
	parse: (xml: string) => XmlObject;
	build: (tree: XmlObject) => string;
	parseDataUrl: (dataUrl: string) => { bytes: Uint8Array; extension: string } | null;
}

function relsPathOf(partPath: string): string {
	const slash = partPath.lastIndexOf('/');
	return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

function nextMediaPath(zip: JSZip, extension: string): string {
	let n = 1;
	while (zip.file(`ppt/media/image${n}.${extension}`)) n += 1;
	return `ppt/media/image${n}.${extension}`;
}

async function ensureDefaultContentType(host: EmbedHost, extension: string): Promise<void> {
	const mime = MIME[extension];
	const xml = await host.zip.file('[Content_Types].xml')?.async('string');
	if (!mime || !xml) return;
	const tree = host.parse(xml);
	const types = (tree['Types'] ?? {}) as XmlObject;
	const defaults = ([] as XmlObject[]).concat((types['Default'] as XmlObject[] | undefined) ?? []);
	if (defaults.some((d) => String(d['@_Extension'] ?? '').toLowerCase() === extension)) return;
	defaults.push({ '@_Extension': extension, '@_ContentType': mime });
	types['Default'] = defaults;
	tree['Types'] = types;
	host.zip.file('[Content_Types].xml', host.build(tree));
}

/**
 * Embed every picture the formats carry without a relationship id: write the
 * media part and the relationship, and return the formats with `relId` set
 * (the input is left alone). Returns the input when there is nothing to embed.
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
	const relsPath = relsPathOf(chartPartPath);
	const relsXml = await host.zip.file(relsPath)?.async('string');
	const tree: XmlObject = relsXml ? host.parse(relsXml) : { Relationships: { '@_xmlns': RELS_NS } };
	const root = (tree['Relationships'] ?? {}) as XmlObject;
	const list = ([] as XmlObject[]).concat((root['Relationship'] as XmlObject[] | undefined) ?? []);
	const used = new Set(list.map((rel) => String(rel['@_Id'])));
	const out: PptxChartElementFormats = { ...formats };
	const bySource = new Map<string, string>();
	for (const [key, record] of pending) {
		const picture = record.fill!.picture!;
		const src = picture.src!;
		let id = bySource.get(src);
		if (!id) {
			const parsed = host.parseDataUrl(src);
			if (!parsed) continue;
			const extension =
				parsed.extension.toLowerCase() === 'jpeg' ? 'jpg' : parsed.extension.toLowerCase();
			const mediaPath = nextMediaPath(host.zip, extension);
			host.zip.file(mediaPath, parsed.bytes);
			await ensureDefaultContentType(host, extension);
			let n = 1;
			while (used.has(`rId${n}`)) n += 1;
			id = `rId${n}`;
			used.add(id);
			list.push({
				'@_Id': id,
				'@_Type': IMAGE_REL_TYPE,
				'@_Target': `../${mediaPath.slice('ppt/'.length)}`,
			});
			bySource.set(src, id);
		}
		out[key] = { ...record, fill: { ...record.fill!, picture: { ...picture, relId: id } } };
	}
	root['Relationship'] = list.length === 1 ? list[0] : list;
	tree['Relationships'] = root;
	host.zip.file(relsPath, host.build(tree));
	return out;
}

/** Declare the relationships namespace on `chartSpace` so `r:embed` resolves. */
export function ensureRelationshipNamespace(chartSpace: XmlObject): void {
	if (!chartSpace['@_xmlns:r']) {
		chartSpace['@_xmlns:r'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	}
}
