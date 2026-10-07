import { chartColorStyleId } from '../../chart/color-style';
import type { SourceIndex } from './package';

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
