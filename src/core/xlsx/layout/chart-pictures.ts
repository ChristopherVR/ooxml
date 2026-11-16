// The pictures of a spreadsheet chart's picture fills, for the painter: a picture chosen in the
// Format pane this session is in the chart model's `pictureSources` under a pending id; one the
// workbook was opened with is a media part its chart part's relationships name, read from the
// source package and cached per package.
import { chartPictureDataUrl } from '../../chart/picture-parts';
import type { ChartPictureSource } from '../../chart/render/chart-picture-paint';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from '../../opc/relationships';
import { RELATIONSHIP_TYPES } from '../../opc/relationship-types';
import type { ChartObject, Workbook } from '../model';

const cache = new WeakMap<Map<string, Uint8Array>, Map<string, string | null>>();
const decoder = new TextDecoder();

function packagePicture(
	parts: Map<string, Uint8Array>,
	chartPart: string,
	relId: string,
): string | undefined {
	let known = cache.get(parts);
	if (!known) cache.set(parts, (known = new Map()));
	const key = `${chartPart}#${relId}`;
	const hit = known.get(key);
	if (hit !== undefined) return hit ?? undefined;
	const rels = parts.get(relationshipsPartFor(chartPart));
	const rel = rels && parseRelationships(decoder.decode(rels)).get(relId);
	let url: string | undefined;
	if (rel && rel.mode !== 'External' && rel.type === RELATIONSHIP_TYPES.image) {
		const part = resolvePartPath(chartPart, rel.target);
		const bytes = parts.get(part);
		url = bytes && chartPictureDataUrl(part, bytes);
	}
	known.set(key, url ?? null);
	return url;
}

/** Resolves the picture fills of `chart` in `workbook` to URLs. */
export function xlsxChartPictureSource(workbook: Workbook, chart: ChartObject): ChartPictureSource {
	const pending = chart.chartSpace?.pictureSources;
	const parts = workbook.source?.parts;
	return (relId) =>
		pending?.[relId] ??
		(parts && chart.partName ? packagePicture(parts, chart.partName, relId) : undefined);
}
