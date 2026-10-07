import { chartColorStyleId } from '../../chart/color-style';
import type { SourceIndex } from './package';
import { readChartStyle } from '../../chart/read-style';

export function readChartPaletteId(
	source: SourceIndex | undefined,
	part: string,
): number | undefined {
	const rel = [...(source?.rels(part).values() ?? [])].find((rel) =>
		rel.type.endsWith('/chartColorStyle'),
	);
	const target = rel && source?.target(part, rel);
	const xml = target && source?.text(target);
	return xml ? chartColorStyleId(xml) : undefined;
}

export function readChartStylePart(
	source: SourceIndex,
	part: string,
	warn: (message: string) => void,
) {
	const rel = [...source.rels(part).values()].find((rel) => rel.type.endsWith('/chartStyle'));
	const target = rel && source.target(part, rel);
	const xml = target && source.text(target);
	if (!rel) return undefined;
	try {
		const style = xml ? readChartStyle(xml) : undefined;
		if (style) return style;
	} catch {
		/* Retain the readable chart and preserve the unsupported style part. */
	}
	warn(`Chart style for ${part} could not be read; its part is kept but formatting is not shown.`);
	return undefined;
}
