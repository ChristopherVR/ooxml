import type { PptxNativeAnimation, XmlObject } from '../types';
import {
	appendContainer,
	findMainSequenceCTn,
	indexEffectNodes,
	maxTimeNodeId,
	removeEffectNode,
	type EffectNodeRef,
} from './animation-timing-tree';
import { buildClickGroupNode, markGroupAutoStart } from './animation-write-sequence-builders';
import { readTimingAttr } from './native-animation-extended-helpers';
import { ensureArray, isXmlObject } from './native-animation-helpers';

function simpleCondition(node: XmlObject | undefined): XmlObject | undefined {
	const cTn = node?.['p:cTn'];
	if (!isXmlObject(cTn) || !isXmlObject(cTn['p:stCondLst'])) return undefined;
	const conditions = ensureArray(cTn['p:stCondLst']['p:cond']);
	const condition = conditions[0];
	return conditions.length === 1 &&
		condition &&
		Object.keys(condition).every((key) => key === '@_delay')
		? condition
		: undefined;
}

/** Move a plain main-sequence effect within PowerPoint's click and time-offset groups. */
export function changeNativeAnimationStart(
	rawTiming: XmlObject,
	ref: EffectNodeRef,
	animations: readonly PptxNativeAnimation[],
	start: 'onClick' | 'withPrevious' | 'afterPrevious',
): boolean {
	const main = findMainSequenceCTn(rawTiming);
	const holder = main?.['p:childTnLst'];
	if (!main || !isXmlObject(holder)) return false;
	const group = ref.chain.find((link) => link.holder === holder);
	const leaf = ref.chain.at(-1);
	const wrapper = ref.chain.at(-2);
	if (
		!group ||
		!leaf ||
		!wrapper ||
		wrapper === group ||
		leaf.key !== 'p:par' ||
		wrapper.key !== 'p:par' ||
		!simpleCondition(leaf.node) ||
		!simpleCondition(wrapper.node)
	)
		return false;
	const groupTime = group.node['p:cTn'];
	if (!isXmlObject(groupTime) || !isXmlObject(groupTime['p:childTnLst'])) return false;
	const groupChildren = groupTime['p:childTnLst'];
	if (wrapper.holder !== groupChildren) return false;
	const refs = indexEffectNodes(rawTiming).filter((item) =>
		item.chain.some((link) => link.holder === holder),
	);
	const index = refs.findIndex((item) => item.cTn === ref.cTn);
	if (index < 0) return false;
	const next = refs[index + 1];
	const previous = refs[index - 1];
	const nextInGroup = next?.chain.some((link) => link.node === group.node) ?? false;
	const previousInGroup = previous?.chain.some((link) => link.node === group.node) ?? false;
	let nextId = maxTimeNodeId(rawTiming) + 1;
	const allocateId = () => nextId++;
	const delay = (value: number | 'indefinite') => ({ 'p:cond': { '@_delay': String(value) } });
	const wrap = (duration: number) => ({
		'p:cTn': {
			'@_id': String(allocateId()),
			'@_fill': 'hold',
			'p:stCondLst': delay(duration),
			'p:childTnLst': { 'p:par': leaf.node },
		},
	});
	if (start === 'onClick') {
		if (!previousInGroup) {
			groupTime['p:stCondLst'] = delay('indefinite');
			return true;
		}
		if (nextInGroup) return false;
		removeEffectNode(ref);
		const groups = ensureArray(holder['p:par']);
		groups.splice(groups.indexOf(group.node) + 1, 0, buildClickGroupNode([wrap(0)], allocateId));
		holder['p:par'] = groups;
		return true;
	}
	if (!previous) {
		const mainId = Number(main['@_id']);
		if (!Number.isSafeInteger(mainId)) return false;
		markGroupAutoStart(group.node, mainId);
		return true;
	}
	if (!previousInGroup && nextInGroup) return false;
	const prevWrapper = previous.chain.at(-2);
	const prevLeaf = previous.chain.at(-1);
	const prevDelay = readTimingAttr(simpleCondition(prevWrapper?.node)?.['@_delay']);
	if (
		!prevWrapper ||
		!prevLeaf ||
		prevDelay === undefined ||
		!simpleCondition(prevLeaf.node) ||
		prevLeaf.key !== 'p:par'
	)
		return false;
	if (start === 'withPrevious') {
		if (prevLeaf.holder === leaf.holder) return true;
		removeEffectNode(ref);
		appendContainer(prevLeaf.holder, 'p:par', leaf.node);
		return true;
	}
	const end = animations.find((animation) => animation.nodeId === Number(previous.cTn['@_id']));
	if (end?.durationMs === undefined || end.repeatCount || end.autoReverse) return false;
	removeEffectNode(ref);
	const wrappers = ensureArray(prevWrapper.holder['p:par']);
	wrappers.splice(
		wrappers.indexOf(prevWrapper.node) + 1,
		0,
		wrap(prevDelay + (end.delayMs ?? 0) + end.durationMs),
	);
	prevWrapper.holder['p:par'] = wrappers;
	return true;
}
