import type { PptxChartHistogramOptions, XmlObject } from '../types';
import type { XmlLookupLike } from './chart-cx-parser';
import { hasLocalName } from './chart-xml-presence';

function parseDoubleOrAutomatic(value: unknown): number | 'auto' | undefined {
	if (value === 'auto') {
		return 'auto';
	}
	const parsed = Number.parseFloat(String(value ?? ''));
	return Number.isFinite(parsed) ? parsed : undefined;
}

/**
 * A `cx:binSize` / `cx:binCount` value. PowerPoint writes `<cx:binCount val="4"/>` and refuses
 * the text form `<cx:binCount>4</cx:binCount>`; the text form is still read, for decks written
 * before that was known.
 */
function binningValue(
	binning: XmlObject | undefined,
	localName: string,
	xmlLookup: XmlLookupLike,
): string | number | boolean | undefined {
	const child = xmlLookup.getChildByLocalName(binning, localName);
	const attribute = child?.['@_val'];
	if (typeof attribute === 'string' || typeof attribute === 'number') {
		return attribute;
	}
	return xmlLookup.getScalarChildByLocalName(binning, localName);
}

/** Parse schema-defined ChartEx histogram and Pareto series properties. */
export function parseCxHistogramOptions(
	series: XmlObject,
	xmlLookup: XmlLookupLike,
): PptxChartHistogramOptions | undefined {
	const layoutId = series['@_layoutId'];
	if (layoutId !== 'clusteredColumn' && layoutId !== 'paretoLine') {
		return undefined;
	}
	const layoutPr = xmlLookup.getChildByLocalName(series, 'layoutPr');
	const binning = xmlLookup.getChildByLocalName(layoutPr, 'binning');
	// A bare `<cx:aggregation/>` reads as an empty string, so test the key (chartex parity harness).
	const aggregation = layoutPr && hasLocalName(layoutPr, 'aggregation') ? true : undefined;
	const rawBinSize = binningValue(binning, 'binSize', xmlLookup);
	const rawBinCount = binningValue(binning, 'binCount', xmlLookup);
	const binSize = Number.parseFloat(String(rawBinSize ?? ''));
	const binCount = Number.parseInt(String(rawBinCount ?? ''), 10);
	const intervalClosed = binning?.['@_intervalClosed'];
	return {
		layout: layoutId === 'paretoLine' ? 'pareto' : 'histogram',
		...(Number.isFinite(binSize) ? { binSize } : {}),
		...(Number.isSafeInteger(binCount) && binCount >= 0 ? { binCount } : {}),
		...(intervalClosed === 'l' || intervalClosed === 'r' ? { intervalClosed } : {}),
		...(binning?.['@_underflow'] !== undefined
			? { underflow: parseDoubleOrAutomatic(binning['@_underflow']) }
			: {}),
		...(binning?.['@_overflow'] !== undefined
			? { overflow: parseDoubleOrAutomatic(binning['@_overflow']) }
			: {}),
		...(aggregation ? { aggregateByCategory: true } : {}),
	};
}
