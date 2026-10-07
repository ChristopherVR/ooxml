import {
	buildChartGradientDef,
	chartGradientMarkup,
	type ChartGradientFill,
} from 'ooxml-core/chart';

/** Shared chart paint supplies SVG thumbnails for Office gradient galleries. */
export function gradientGalleryPreview(id: string, fill: ChartGradientFill): string {
	const def = buildChartGradientDef(id, fill, { width: 38, height: 38, shape: 'rect' });
	return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 40"><defs>${chartGradientMarkup(def)}</defs><rect x="1" y="1" width="38" height="38" fill="url(#${def.id})" stroke="#999"/></svg>`;
}
