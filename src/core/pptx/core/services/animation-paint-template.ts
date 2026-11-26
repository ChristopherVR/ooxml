import type { PptxElementAnimation, XmlObject } from '../types';
import { indexEffectNodes } from './animation-timing-tree';
import { isXmlObject } from './native-animation-helpers';

/** Capture an imported effect's local behaviours, excluding slide-bound triggers and sounds. */
export function captureNativeAnimationTemplate(
	rawTiming: XmlObject,
	nodeId: number,
): PptxElementAnimation['nativePaintTemplate'] {
	const matches = indexEffectNodes(rawTiming).filter((ref) => Number(ref.cTn['@_id']) === nodeId);
	const ref = matches[0];
	if (matches.length !== 1 || !ref?.spid) return undefined;
	const children = ref.cTn['p:childTnLst'];
	if (!isXmlObject(children)) return undefined;
	return { sourceShapeId: ref.spid, children, presetClass: ref.presetClass };
}

/** Clone and rebind a local behaviour tree to the destination shape and timing ID space. */
export function buildPaintedAnimationChildren(
	template: PptxElementAnimation['nativePaintTemplate'],
	presetClass: string,
	shapeId: string,
	allocateId: () => number,
): XmlObject | undefined {
	if (template?.presetClass !== presetClass) return undefined;
	const children = structuredClone(template.children);
	const queue: Array<readonly [string, unknown]> = Object.entries(children);
	for (let i = 0; i < queue.length; i++) {
		if (queue.length > 4096) return undefined;
		const entry = queue[i];
		if (!entry) continue;
		const [key, node] = entry;
		if (Array.isArray(node)) {
			if (queue.length + node.length > 4096) return undefined;
			for (const child of node) queue.push([key, child]);
			continue;
		}
		if (key.startsWith('@_r:') || key === 'p:tn' || key === 'p:rtn' || key === 'p:subSp')
			return undefined;
		if (!isXmlObject(node)) continue;
		if (key === 'p:cTn') node['@_id'] = String(allocateId());
		if (key === 'p:spTgt') {
			if (String(node['@_spid']) !== template.sourceShapeId) return undefined;
			node['@_spid'] = shapeId;
		}
		for (const child of Object.entries(node)) queue.push(child);
	}
	return children;
}
