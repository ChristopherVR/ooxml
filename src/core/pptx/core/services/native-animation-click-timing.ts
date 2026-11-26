import type { PptxNativeAnimation, XmlObject } from '../types';
import { resolveEffectTiming } from './animation-effect-timing';
import { countTimeNodeChildren, type EffectNodeRef } from './animation-timing-tree';
import { readTimingAttr } from './native-animation-extended-helpers';
import { ensureArray, isXmlObject } from './native-animation-helpers';

function timingNode(node: XmlObject | undefined) {
	const cTn = node?.['p:cTn'];
	return isXmlObject(cTn) ? cTn : undefined;
}

export function simpleCondition(node: XmlObject | undefined): XmlObject | undefined {
	const cTn = timingNode(node);
	if (!isXmlObject(cTn?.['p:stCondLst'])) return undefined;
	const conditions = ensureArray(cTn['p:stCondLst']['p:cond']);
	const condition = conditions[0];
	return conditions.length === 1 &&
		condition &&
		Object.keys(condition).every((key) => key === '@_delay')
		? condition
		: undefined;
}

/** The direct effect wrapper inside a main-sequence click group. */
export function mainEffectWrapper(ref: EffectNodeRef) {
	const [main, group, wrapper, leaf] = ref.chain.slice(-4);
	if (
		timingNode(main?.node)?.['@_nodeType'] !== 'mainSeq' ||
		group?.key !== 'p:par' ||
		wrapper?.key !== 'p:par' ||
		leaf?.key !== 'p:par'
	)
		return undefined;
	return wrapper;
}

/** End offset relative to the wrapper start, including overlapping siblings. */
export function effectWrapperSpan(
	wrapper: XmlObject,
	animations: ReadonlyMap<number, PptxNativeAnimation>,
	exclude?: ReadonlySet<XmlObject>,
): number | undefined {
	const children = timingNode(wrapper)?.['p:childTnLst'];
	if (!isXmlObject(children)) return undefined;
	const effects = ensureArray(children['p:par']);
	if (effects.length !== countTimeNodeChildren(children) || effects.length === 0) return undefined;
	let span = 0;
	for (const effect of effects) {
		if (exclude?.has(effect)) continue;
		const anim = animations.get(Number(timingNode(effect)?.['@_id']));
		if (
			anim?.durationMs === undefined ||
			anim.repeatDurMs ||
			anim.iterate ||
			anim.endConditions?.length ||
			!simpleCondition(effect)
		)
			return undefined;
		const { activeDurationMs, iterationCount } = resolveEffectTiming(anim, anim.durationMs);
		if (!Number.isFinite(iterationCount)) return undefined;
		span = Math.max(span, (anim.delayMs ?? 0) + Math.round(activeDurationMs));
	}
	return span;
}

/** Recompute absolute offsets for the dependent After Previous wrappers. */
export function reflowAfterPrevious(
	holder: XmlObject,
	from: XmlObject,
	animations: ReadonlyMap<number, PptxNativeAnimation>,
): void {
	const wrappers = ensureArray(holder['p:par']);
	const index = wrappers.indexOf(from);
	if (index < 0) return;
	let previous = from;
	for (const wrapper of wrappers.slice(index + 1)) {
		const children = timingNode(wrapper)?.['p:childTnLst'];
		if (!isXmlObject(children)) return;
		const first = ensureArray(children['p:par'])[0];
		const nodeType = timingNode(first)?.['@_nodeType'];
		if (nodeType !== 'afterEffect' && nodeType !== 'afterPrevious' && nodeType !== 'afterPrev')
			return;
		const start = readTimingAttr(simpleCondition(previous)?.['@_delay']);
		const span = effectWrapperSpan(previous, animations);
		const condition = simpleCondition(wrapper);
		if (start === undefined || span === undefined || !condition) return;
		condition['@_delay'] = String(start + span);
		previous = wrapper;
	}
}
