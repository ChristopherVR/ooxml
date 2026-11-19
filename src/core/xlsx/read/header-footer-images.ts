// Header and footer pictures (`&G`): the sheet's `legacyDrawingHF` is a VML drawing whose shapes
// are named `LH`, `CH`, `RH`, `LF`, `CF`, `RF` (with `FIRST` or `EVEN` for those pages) and point
// at image parts through `o:relid`. Read for display only; the VML part is kept verbatim.
import { relAttr, type XmlElement } from '../../xml/index';
import { xmlTags } from '../../xml/tags';
import type { HeaderFooterImage } from '../model';
import type { SourceIndex } from './package';

const UNIT_PT: Record<string, number> = { pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, px: 0.75 };

/** A CSS length from a VML `style` (`width:30pt`) in points. */
function styleLength(style: string, property: string): number {
	const match = new RegExp(
		`(?:^|[;\\s])${property}\\s*:\\s*([\\d.]+)\\s*(pt|in|cm|mm|px)?`,
		'i',
	).exec(style);
	return match ? Number(match[1]) * (UNIT_PT[(match[2] ?? 'pt').toLowerCase()] ?? 1) : 0;
}

/** Each `v:shape` element's attributes and its first `v:imagedata`'s, in document order. */
function* vmlShapes(vml: string): Generator<{ attributes: string; data: string | undefined }> {
	let shape: { attributes: string; data: string | undefined } | undefined;
	for (const tag of xmlTags(vml)) {
		if (shape === undefined) {
			if (!tag.closing && !tag.selfClosing && tag.name === 'v:shape')
				shape = { attributes: tag.attributes, data: undefined };
		} else if (tag.closing && tag.name === 'v:shape') {
			yield shape;
			shape = undefined;
		} else if (!tag.closing && tag.name === 'v:imagedata') shape.data ??= tag.attributes;
	}
}

export function readHeaderFooterImages(
	source: SourceIndex,
	sheetPart: string,
	legacyHF: XmlElement | undefined,
): Record<string, HeaderFooterImage> | undefined {
	if (!legacyHF) return undefined;
	const rel = source.rels(sheetPart).get(relAttr(legacyHF, 'id') ?? '');
	const vmlPart = rel ? source.target(sheetPart, rel) : undefined;
	const vml = vmlPart ? source.text(vmlPart) : undefined;
	if (!vmlPart || !vml) return undefined;
	const rels = source.rels(vmlPart);
	const images: Record<string, HeaderFooterImage> = {};
	for (const { attributes, data } of vmlShapes(vml)) {
		const id = /\bid="([^"]+)"/.exec(attributes)?.[1];
		const style = /style=(?:'([^']*)'|"([^"]*)")/.exec(attributes);
		const relId = /\bo:relid="([^"]+)"/.exec(data ?? '')?.[1];
		const imageRel = relId ? rels.get(relId) : undefined;
		const partName = imageRel ? source.target(vmlPart, imageRel) : undefined;
		if (!id || !partName || !source.has(partName)) continue;
		const css = style?.[1] ?? style?.[2] ?? '';
		const title = /\bo:title="([^"]*)"/.exec(data ?? '')?.[1];
		images[id] = {
			partName,
			contentType: source.contentType(partName) ?? 'application/octet-stream',
			widthPt: styleLength(css, 'width'),
			heightPt: styleLength(css, 'height'),
			...(title ? { title } : {}),
		};
	}
	return Object.keys(images).length ? images : undefined;
}
