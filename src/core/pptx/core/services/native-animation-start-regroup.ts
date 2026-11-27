/**
 * Regrouping of a main-sequence timing tree when an effect's Start changes.
 *
 * PowerPoint keeps the main sequence as click groups (`p:par` with an
 * `indefinite` start) holding time-offset wrappers (`p:par` with an absolute
 * delay from the click) holding the effects. Changing Start moves effects
 * between those levels:
 *
 * - On Click on an effect in the middle of a click group splits the group: the
 *   effect, the later With Previous effects of its wrapper and every later
 *   wrapper of the group move to a new click group after it. The moved wrappers
 *   keep their offsets relative to the effect, then After Previous offsets are
 *   recomputed from the measured wrapper spans.
 * - With Previous or After Previous on the first effect of a click group merges
 *   the group into the previous one: its later wrappers follow the effect into
 *   the previous group, rebased onto the new start of the effect.
 *
 * The rules only touch plain `p:par` containers with single delay conditions;
 * the caller refuses everything else so no timing is ever guessed.
 *
 * @module services/native-animation-start-regroup
 */
import type { XmlObject } from '../types';
import {
	countTimeNodeChildren,
	removeEffectNode,
	type EffectNodeRef,
} from './animation-timing-tree';
import { simpleCondition } from './native-animation-click-timing';
import { readTimingAttr } from './native-animation-extended-helpers';
import { ensureArray, isXmlObject } from './native-animation-helpers';

/** Why a Start change was refused instead of regrouping the timing tree. */
export type NativeAnimationStartRefusal =
	/** Triggered (interactive) or event-conditioned timing: there is no click group to move to. */
	| 'eventTiming'
	/** The effect is not inside the plain click group, wrapper, effect nesting. */
	| 'nestedTiming'
	/** A wrapper or effect that would move has no plain numeric delay. */
	| 'conditionTiming'
	/** After Previous needs the end of an effect whose length cannot be measured. */
	| 'unmeasuredTiming';

/** The plain `p:par` children of a container, or `undefined` when it holds anything else. */
export function plainParChildren(node: XmlObject | undefined): XmlObject[] | undefined {
	const cTn = node?.['p:cTn'];
	const children = isXmlObject(cTn) ? cTn['p:childTnLst'] : undefined;
	if (!isXmlObject(children)) return undefined;
	const pars = ensureArray(children['p:par']);
	return pars.length > 0 && pars.length === countTimeNodeChildren(children) ? pars : undefined;
}

/** The numeric start offset of a wrapper, or `undefined` for event or indefinite starts. */
export function wrapperOffset(wrapper: XmlObject): number | undefined {
	return readTimingAttr(simpleCondition(wrapper)?.['@_delay']);
}

export function setWrapperOffset(wrapper: XmlObject, offset: number): void {
	const cTn = wrapper['p:cTn'];
	if (isXmlObject(cTn)) cTn['p:stCondLst'] = { 'p:cond': { '@_delay': String(offset) } };
}

function setChildren(holder: XmlObject, nodes: XmlObject[]): void {
	if (nodes.length === 0) delete holder['p:par'];
	else holder['p:par'] = nodes.length === 1 ? nodes[0]! : nodes;
}

export function childList(node: XmlObject): XmlObject | undefined {
	const cTn = node['p:cTn'];
	const children = isXmlObject(cTn) ? cTn['p:childTnLst'] : undefined;
	return isXmlObject(children) ? children : undefined;
}

/** Wrappers after `wrapper` in its click group, checked for plain numeric offsets. */
export function followingWrappers(
	group: XmlObject,
	wrapper: XmlObject,
): XmlObject[] | NativeAnimationStartRefusal {
	const wrappers = plainParChildren(group);
	if (!wrappers) return 'nestedTiming';
	const index = wrappers.indexOf(wrapper);
	if (index < 0) return 'nestedTiming';
	const after = wrappers.slice(index + 1);
	if (after.some((node) => wrapperOffset(node) === undefined || !plainParChildren(node)))
		return 'conditionTiming';
	return after;
}

/** Move `wrappers` out of their click group, rebased by `shift` milliseconds. */
export function detachWrappers(group: XmlObject, wrappers: readonly XmlObject[], shift: number) {
	const children = childList(group);
	if (!children) return;
	const moving = new Set(wrappers);
	setChildren(
		children,
		ensureArray(children['p:par']).filter((node) => !moving.has(node)),
	);
	for (const wrapper of wrappers) {
		setWrapperOffset(wrapper, Math.max(0, (wrapperOffset(wrapper) ?? 0) + shift));
	}
}

/** Insert `nodes` after `anchor` in `holder` (at the end when the anchor is gone). */
export function insertAfter(holder: XmlObject, anchor: XmlObject, nodes: readonly XmlObject[]) {
	const siblings = ensureArray(holder['p:par']);
	const index = siblings.indexOf(anchor);
	siblings.splice(index < 0 ? siblings.length : index + 1, 0, ...nodes);
	setChildren(holder, siblings);
}

/**
 * Take the effect at `ref` and the later effects of its wrapper out of the tree.
 *
 * When the effect opens its wrapper the wrapper itself is returned, detached
 * and unchanged, so its id and attributes survive; otherwise the effects are
 * removed and the caller wraps them in a new wrapper. An emptied click group is
 * pruned by `removeEffectNode`, never an unrelated one.
 */
export function takeWrapperTail(
	ref: EffectNodeRef,
	tail: readonly EffectNodeRef[],
): { wrapper: XmlObject } | { effects: XmlObject[] } {
	const leaf = ref.chain.at(-1)!;
	const wrapper = ref.chain.at(-2)!;
	const effects = ensureArray(leaf.holder['p:par']);
	if (effects[0] === leaf.node && effects.length === tail.length) {
		const group = ref.chain.at(-3);
		const children = group ? childList(group.node) : undefined;
		if (children) {
			setChildren(
				children,
				ensureArray(children['p:par']).filter((node) => node !== wrapper.node),
			);
			return { wrapper: wrapper.node };
		}
	}
	const nodes = tail.map((item) => item.chain.at(-1)!.node);
	for (const item of tail) removeEffectNode(item);
	return { effects: nodes };
}

/** Whether a time node starts on an event (a trigger) rather than a delay. */
export function hasEventCondition(node: XmlObject | undefined): boolean {
	const cTn = node?.['p:cTn'];
	const list = isXmlObject(cTn) ? cTn['p:stCondLst'] : undefined;
	return (
		isXmlObject(list) &&
		ensureArray(list['p:cond']).some((cond) => cond['@_evt'] !== undefined || cond['p:tgtEl'])
	);
}
