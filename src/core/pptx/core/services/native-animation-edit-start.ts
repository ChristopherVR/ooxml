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
import {
	effectWrapperSpan,
	reflowAfterPrevious,
	simpleCondition,
} from './native-animation-click-timing';
import { isXmlObject } from './native-animation-helpers';
import {
	childList,
	detachWrappers,
	followingWrappers,
	hasEventCondition,
	insertAfter,
	plainParChildren,
	setWrapperOffset,
	takeWrapperTail,
	wrapperOffset,
	type NativeAnimationStartRefusal,
} from './native-animation-start-regroup';

export type NativeAnimationStart = 'onClick' | 'withPrevious' | 'afterPrevious';
export type { NativeAnimationStartRefusal } from './native-animation-start-regroup';
/** `'ok'` when the Start change applies (or would apply, for a dry run). */
export type NativeAnimationStartResult = 'ok' | NativeAnimationStartRefusal;

/**
 * Move a main-sequence effect between PowerPoint's click and time-offset groups.
 *
 * Splitting a click group (On Click in its middle) and merging one into the
 * previous group (With or After Previous on its first effect) regroup the tree
 * as described in `native-animation-start-regroup`. Trees that are not plain
 * click group, wrapper, effect nesting with numeric delays are refused with the
 * reason, and a dry run returns the same answer as the edit without changing
 * anything.
 */
export function changeNativeAnimationStart(
	rawTiming: XmlObject,
	ref: EffectNodeRef,
	animations: readonly PptxNativeAnimation[],
	start: NativeAnimationStart,
	dryRun = false,
): NativeAnimationStartResult {
	const main = findMainSequenceCTn(rawTiming);
	const holder = main?.['p:childTnLst'];
	if (!main || !isXmlObject(holder)) return 'eventTiming';
	const group = ref.chain.find((link) => link.holder === holder);
	if (!group) return 'eventTiming';
	const leaf = ref.chain.at(-1);
	const wrapper = ref.chain.at(-2);
	if (
		!leaf ||
		!wrapper ||
		wrapper === group ||
		group.key !== 'p:par' ||
		leaf.key !== 'p:par' ||
		wrapper.key !== 'p:par' ||
		wrapper.holder !== childList(group.node) ||
		!plainParChildren(wrapper.node)
	)
		return 'nestedTiming';
	if (hasEventCondition(leaf.node)) return 'eventTiming';
	const ownOffset = wrapperOffset(wrapper.node);
	if (!simpleCondition(leaf.node) || ownOffset === undefined) return 'conditionTiming';
	const refs = indexEffectNodes(rawTiming).filter((item) =>
		item.chain.some((link) => link.holder === holder),
	);
	const index = refs.findIndex((item) => item.cTn === ref.cTn);
	if (index < 0) return 'nestedTiming';
	const previous = refs[index - 1];
	const previousInGroup = previous?.chain.some((link) => link.node === group.node) ?? false;
	const tail = refs.slice(index).filter((item) => item.chain.at(-1)?.holder === leaf.holder);
	const tailNodes = tail.map((item) => item.chain.at(-1)?.node ?? leaf.node);
	if (tailNodes.some((node) => !simpleCondition(node))) return 'conditionTiming';
	let nextId: number | undefined;
	const allocateId = () => {
		nextId ??= maxTimeNodeId(rawTiming) + 1;
		return nextId++;
	};
	const wrap = (offset: number, effects: XmlObject[]): XmlObject => ({
		'p:cTn': {
			'@_id': String(allocateId()),
			'@_fill': 'hold',
			'p:stCondLst': { 'p:cond': { '@_delay': String(offset) } },
			'p:childTnLst': { 'p:par': effects.length === 1 ? effects[0] : effects },
		},
	});
	const byId = new Map(
		animations.flatMap((animation) =>
			animation.nodeId === undefined ? [] : [[animation.nodeId, animation] as const],
		),
	);
	const groupTime = group.node['p:cTn'];
	if (!isXmlObject(groupTime)) return 'nestedTiming';

	if (start === 'onClick') {
		if (!previousInGroup) {
			if (!dryRun) groupTime['p:stCondLst'] = { 'p:cond': { '@_delay': 'indefinite' } };
			return 'ok';
		}
		// Split: the effect and everything after it in its click group start a new click.
		const followers = followingWrappers(group.node, wrapper.node);
		if (typeof followers === 'string') return followers;
		if (dryRun) return 'ok';
		detachWrappers(group.node, followers, -ownOffset);
		const taken = takeWrapperTail(ref, tail);
		const head = 'wrapper' in taken ? taken.wrapper : wrap(0, taken.effects);
		setWrapperOffset(head, 0);
		const clickGroup = buildClickGroupNode([head, ...followers], allocateId);
		insertAfter(holder, group.node, [clickGroup]);
		const children = childList(clickGroup);
		if (children) reflowAfterPrevious(children, head, byId);
		return 'ok';
	}

	if (!previous) {
		const mainId = Number(main['@_id']);
		if (!Number.isSafeInteger(mainId)) return 'conditionTiming';
		if (!dryRun) markGroupAutoStart(group.node, mainId);
		return 'ok';
	}
	const prevWrapper = previous.chain.at(-2);
	const prevLeaf = previous.chain.at(-1);
	const prevGroup = previous.chain.find((link) => link.holder === holder);
	if (
		!prevWrapper ||
		!prevLeaf ||
		!prevGroup ||
		prevLeaf.key !== 'p:par' ||
		prevWrapper.holder !== childList(prevGroup.node)
	)
		return 'nestedTiming';
	const prevDelay = wrapperOffset(prevWrapper.node);
	if (prevDelay === undefined || !simpleCondition(prevLeaf.node)) return 'conditionTiming';
	// Merge: the first effect of a click group joins the previous group, and the
	// rest of its group follows it, rebased onto the effect's new start.
	const followers = previousInGroup ? [] : followingWrappers(group.node, wrapper.node);
	if (typeof followers === 'string') return followers;

	if (start === 'withPrevious') {
		if (dryRun || prevLeaf.holder === leaf.holder) return 'ok';
		detachWrappers(group.node, followers, prevDelay - ownOffset);
		for (const item of tail) removeEffectNode(item);
		for (const node of tailNodes) appendContainer(prevLeaf.holder, 'p:par', node);
		insertAfter(prevWrapper.holder, prevWrapper.node, followers);
		reflowAfterPrevious(prevWrapper.holder, prevWrapper.node, byId);
		return 'ok';
	}
	const span = effectWrapperSpan(prevWrapper.node, byId, new Set(tailNodes));
	if (span === undefined) return 'unmeasuredTiming';
	if (dryRun) return 'ok';
	const begin = prevDelay + span;
	detachWrappers(group.node, followers, begin - ownOffset);
	for (const item of tail) removeEffectNode(item);
	const nextWrapper = wrap(begin, tailNodes);
	insertAfter(prevWrapper.holder, prevWrapper.node, [nextWrapper, ...followers]);
	ref.cTn['@_nodeType'] = 'afterEffect';
	reflowAfterPrevious(prevWrapper.holder, prevWrapper.node, byId);
	return 'ok';
}
