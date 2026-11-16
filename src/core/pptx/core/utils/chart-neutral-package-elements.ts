/**
 * The chart part elements PowerPoint's own codecs encode for a neutral chart save
 * (`chart-neutral-save-part-settings.ts`): the print settings (`applyChartPrintSettings`) and the
 * pivot chart formats (`applyChartPivotFormats`), each merged with the authored element as the
 * object-tree save merges it, then handed to the shared writer as element XML. A copy of the
 * parsed part is encoded, so the object-tree save that may follow starts from the source.
 *
 * @module utils/chart-neutral-package-elements
 */

import type { PptxChartData, XmlObject } from '../types';
import type { ChartColorParser } from './chart-pivot-format-fields';
import { applyChartPivotFormats } from './chart-pivot-formats';
import { applyChartPrintSettings } from './chart-print-settings';
import { same } from './chart-neutral-save-common';
import type { NeutralChartPackageElements } from './chart-neutral-save';

const NAMESPACES = {
	c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

export interface ChartElementCodecContext {
	localName: (key: string) => string;
	build: (tree: XmlObject) => string;
	parseColor: ChartColorParser['parseColor'];
}

/** The element `local` of `parent`, built standalone (namespaces declared), or `null` when absent. */
function elementXml(
	parent: XmlObject,
	local: string,
	context: ChartElementCodecContext,
): string | null {
	const key = Object.keys(parent).find((name) => context.localName(name) === local);
	if (!key) return null;
	const node = parent[key] as XmlObject;
	const prefix = key.includes(':') ? key.slice(0, key.indexOf(':')) : 'c';
	const declarations: XmlObject = {};
	for (const [name, uri] of Object.entries(NAMESPACES))
		declarations[name === 'c' ? `@_xmlns:${prefix}` : `@_xmlns:${name}`] = uri;
	return context
		.build({ [key]: { ...declarations, ...node } as XmlObject } as XmlObject)
		.replace(/^<\?xml[^>]*\?>\s*/u, '');
}

/**
 * The elements `after` changed from `before` that the neutral save needs encoded, from the parsed
 * part's `chartSpace` node.
 */
export function encodeNeutralChartElements(
	chartSpace: XmlObject,
	before: PptxChartData,
	after: PptxChartData,
	context: ChartElementCodecContext,
): NeutralChartPackageElements {
	const elements: NeutralChartPackageElements = {};
	if (after.printSettings !== undefined && !same(before.printSettings, after.printSettings)) {
		const space = structuredClone(chartSpace);
		applyChartPrintSettings(space, after.printSettings, context.localName);
		elements.printSettings = elementXml(space, 'printSettings', context);
	}
	if (after.pivotFormats !== undefined && !same(before.pivotFormats, after.pivotFormats)) {
		const chartKey = Object.keys(chartSpace).find((name) => context.localName(name) === 'chart');
		if (chartKey) {
			const chart = structuredClone(chartSpace[chartKey] as XmlObject);
			applyChartPivotFormats(chart, after.pivotFormats, context.localName, {
				parseColor: context.parseColor,
			});
			elements.pivotFormats = elementXml(chart, 'pivotFmts', context);
		}
	}
	return elements;
}
