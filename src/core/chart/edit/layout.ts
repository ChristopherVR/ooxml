// Manual placement edits retain extension children at both layout levels.
import { NS, elements, first, parseXml } from '../../xml';
import type { ChartSpace } from '../model';
import type { ChartLayout } from '../model-series';
import type { ChartManualLayout } from '../manual-layout';
import { CHART_ROOT_BINDINGS, innerXml } from '../xml-fragment';
import { editChartSpace } from './draft';

const FIELDS = [
	['layoutTarget', 'layoutTarget'],
	['xMode', 'xMode'],
	['yMode', 'yMode'],
	['widthMode', 'wMode'],
	['heightMode', 'hMode'],
	['x', 'x'],
	['y', 'y'],
	['width', 'w'],
	['height', 'h'],
] as const;

/** Replace the manual values, keeping opaque source XML. Null returns to automatic placement. */
export function editChartLayout(
	current: ChartLayout | undefined,
	value: ChartManualLayout | null,
): ChartLayout {
	for (const key of ['x', 'y', 'width', 'height'] as const)
		if (value?.[key] !== undefined && !Number.isFinite(value[key]))
			throw new RangeError(`Chart layout ${key} must be finite`);
	const bindings = [...CHART_ROOT_BINDINGS].map(([p, uri]) => ` xmlns:${p}="${uri}"`).join('');
	const root = parseXml(
		`<c:layout${bindings}>${current?.sourceXml ?? ''}</c:layout>`,
	).documentElement;
	let manual = first(root, 'manualLayout', NS.c);
	if (!manual && value) {
		manual = root.ownerDocument.createElementNS(NS.c, 'c:manualLayout');
		root.insertBefore(manual, root.firstChild);
	}
	if (manual) {
		const previous = new Map(FIELDS.map(([, tag]) => [tag, first(manual, tag, NS.c)]));
		for (const [, tag] of FIELDS) {
			const old = first(manual, tag, NS.c);
			if (old) manual.removeChild(old);
		}
		const before = manual.firstChild;
		for (const [key, tag] of FIELDS) {
			if (value?.[key] === undefined) continue;
			const node = previous.get(tag) ?? root.ownerDocument.createElementNS(NS.c, `c:${tag}`);
			node.setAttribute('val', String(value[key]));
			manual.insertBefore(node, before);
		}
		if (!elements(manual).length) root.removeChild(manual);
	}
	return { ...(value ?? {}), sourceXml: innerXml(root) };
}

export function setChartLayout(
	space: ChartSpace,
	region: 'title' | 'legend' | 'plotArea',
	value: ChartManualLayout | null,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const holder = draft[region];
		if (!holder) throw new RangeError(`The chart has no ${region}`);
		holder.layout = editChartLayout(holder.layout, value);
	});
}
