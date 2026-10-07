import { buildChartGradientDef, type ChartGradientFill } from 'ooxml-core/chart';

/** Shared chart paint supplies SVG thumbnails for Office gradient galleries. */
export function gradientGalleryPreview(id: string, fill: ChartGradientFill): string {
	const def = buildChartGradientDef(id, fill);
	const geometry =
		def.kind === 'linearGradient'
			? `x1="${def.x1}" y1="${def.y1}" x2="${def.x2}" y2="${def.y2}"`
			: `cx="${def.cx}" cy="${def.cy}" r="${def.r}"`;
	const stops = def.stops
		.map(
			(stop) =>
				`<stop offset="${stop.offset}" stop-color="${stop.color}" stop-opacity="${stop.opacity ?? 1}"/>`,
		)
		.join('');
	return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 40"><defs><${def.kind} id="${def.id}" ${geometry}>${stops}</${def.kind}></defs><rect x="1" y="1" width="38" height="38" fill="url(#${def.id})" stroke="#999"/></svg>`;
}
