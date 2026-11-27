import type {
	PptxAnimationDirection,
	PptxAnimationPreset,
	PptxNativeAnimation,
	XmlObject,
} from '../types';
import { getCapturedPreset } from './animation-behavior-captured';
import { refreshEffectBehaviors, rescaleBehaviors } from './animation-timing-surgical-behaviors';
import { indexEffectNodes, maxTimeNodeId } from './animation-timing-tree';
import { resolveOoxmlPresetMapping } from './animation-write-mappings';
import { triggerToNodeType } from './animation-write-mappings-motion';
import { mainEffectWrapper, reflowAfterPrevious } from './native-animation-click-timing';
import {
	changeNativeAnimationStart,
	type NativeAnimationStart,
	type NativeAnimationStartRefusal,
} from './native-animation-edit-start';
import { ensureArray, isXmlObject } from './native-animation-helpers';
import { PptxNativeAnimationService } from './PptxNativeAnimationService';

export type NativeEffectTimingEdit =
	| { field: 'duration' | 'delay'; value: number }
	| { field: 'direction'; value: PptxAnimationDirection }
	| { field: 'start'; value: NativeAnimationStart };

function nativeTimingTarget(
	rawTiming: XmlObject,
	animations: readonly PptxNativeAnimation[],
	nodeId: number,
) {
	const matching = animations.filter((animation) => animation.nodeId === nodeId);
	const animation = matching[0];
	if (matching.length !== 1 || !animation?.presetClass || !animation.targetId) return undefined;

	const refs = indexEffectNodes(rawTiming).filter((ref) => Number(ref.cTn['@_id']) === nodeId);
	const ref = refs[0];
	return refs.length === 1 && ref ? { animation, ref } : undefined;
}

/** The Start changes the timing tree supports, and why the others are refused. */
export interface NativeAnimationStartAvailability {
	options: NativeAnimationStart[];
	/** The reason for the first refused Start, when any is refused. */
	refusal?: NativeAnimationStartRefusal | undefined;
}

/** Dry-run every Start change for one effect without altering the tree. */
export function nativeAnimationStartAvailability(
	rawTiming: XmlObject,
	animations: readonly PptxNativeAnimation[],
	nodeId: number,
): NativeAnimationStartAvailability {
	const target = nativeTimingTarget(rawTiming, animations, nodeId);
	if (!target) return { options: [], refusal: 'nestedTiming' };
	if (target.animation.interactiveSequence) return { options: [], refusal: 'eventTiming' };
	const start = target.animation.startMode ?? target.animation.trigger;
	const modes: NativeAnimationStart[] = ['onClick', 'withPrevious', 'afterPrevious'];
	const options: NativeAnimationStart[] = [];
	let refusal: NativeAnimationStartRefusal | undefined;
	for (const mode of modes) {
		if (mode === start) continue;
		const result = changeNativeAnimationStart(rawTiming, target.ref, animations, mode, true);
		if (result === 'ok') options.push(mode);
		else refusal ??= result;
	}
	return { options, refusal };
}

/** Return the Start changes supported by the timing tree without altering it. */
export function nativeAnimationStartOptions(
	rawTiming: XmlObject,
	animations: readonly PptxNativeAnimation[],
	nodeId: number,
): NativeAnimationStart[] {
	return nativeAnimationStartAvailability(rawTiming, animations, nodeId).options;
}

/** Patch the exact native effect and leave sibling effect and behaviour data intact. */
export function editNativeEffectTiming(
	rawTiming: XmlObject,
	animations: readonly PptxNativeAnimation[],
	nodeId: number,
	edit: NativeEffectTimingEdit,
): { rawTiming: XmlObject; nativeAnimations: PptxNativeAnimation[] } | undefined {
	const timing = structuredClone(rawTiming);
	const target = nativeTimingTarget(timing, animations, nodeId);
	if (!target) return undefined;
	const { animation, ref } = target;
	const { cTn } = ref;
	let next: PptxNativeAnimation = { ...animation };
	if (edit.field === 'direction') {
		if (!ref.spid || (animation.presetClass !== 'entr' && animation.presetClass !== 'exit'))
			return undefined;
		const token = `${animation.presetClass}.${animation.presetId}`;
		const mapping = resolveOoxmlPresetMapping(
			token,
			(cls, id) => getCapturedPreset(cls, id)?.defaultSubtype,
		);
		if (!mapping) return undefined;
		let id = maxTimeNodeId(timing) + 1;
		const before = cTn['@_presetSubtype'];
		refreshEffectBehaviors(
			cTn,
			{ elementId: ref.spid, direction: edit.value },
			token as PptxAnimationPreset,
			mapping,
			ref.spid,
			() => id++,
		);
		if (cTn['@_presetSubtype'] === before) return undefined;
		const parsed = new PptxNativeAnimationService()
			.parseNativeAnimations({ 'p:sld': { 'p:timing': timing } })
			?.find((item) => item.nodeId === nodeId);
		if (!parsed) return undefined;
		next = {
			...animation,
			presetSubtype: parsed.presetSubtype,
			behaviors: parsed.behaviors,
			effectFilter: parsed.effectFilter,
			accel: parsed.accel,
			decel: parsed.decel,
		};
	} else if (edit.field === 'start') {
		if (
			animation.interactiveSequence ||
			(animation.startMode ?? animation.trigger) === edit.value ||
			changeNativeAnimationStart(timing, ref, animations, edit.value) !== 'ok'
		)
			return undefined;
		cTn['@_nodeType'] = triggerToNodeType(edit.value);
		next.startMode = edit.value;
		next.trigger = edit.value;
	} else {
		if (!Number.isFinite(edit.value)) return undefined;
		const value = Math.max(
			edit.field === 'duration' ? 1 : 0,
			Math.min(10000, Math.round(edit.value)),
		);
		if (edit.field === 'duration') {
			const current = animation.durationMs;
			if (current === value) return undefined;
			const children = cTn['p:childTnLst'];
			if (current && isXmlObject(children)) {
				rescaleBehaviors(children, current, value);
				next.behaviors = animation.behaviors?.map((behavior) => {
					const { durationMs, delayMs } = behavior.timing;
					const ratio = value / current;
					const closes = durationMs === 1 && delayMs === current - 1;
					return {
						...behavior,
						timing: {
							...behavior.timing,
							durationMs:
								durationMs !== undefined && durationMs > 1
									? Math.max(1, Math.round(durationMs * ratio))
									: durationMs,
							delayMs: closes
								? value - 1
								: delayMs === undefined
									? undefined
									: Math.round(delayMs * ratio),
						},
					};
				});
			}
			cTn['@_dur'] = String(value);
			next.durationMs = value;
		} else {
			const conditions = cTn['p:stCondLst'];
			const cond = isXmlObject(conditions) ? ensureArray(conditions['p:cond']) : [];
			if (cond.length > 1 || value === animation.delayMs) return undefined;
			if (cond[0]) cond[0]['@_delay'] = String(value);
			else cTn['p:stCondLst'] = { 'p:cond': { '@_delay': String(value) } };
			next.delayMs = value;
			next.triggerDelayMs = value;
			next.startConditions = next.startConditions?.map((condition) => ({
				...condition,
				delayMs: value,
			}));
		}
	}
	let nativeAnimations = animations.map((item) => (item === animation ? next : item));
	if (edit.field === 'duration' || edit.field === 'delay') {
		const wrapper = mainEffectWrapper(ref);
		if (wrapper) {
			const byId = new Map(
				nativeAnimations.flatMap((item) =>
					item.nodeId === undefined ? [] : [[item.nodeId, item] as const],
				),
			);
			reflowAfterPrevious(wrapper.holder, wrapper.node, byId);
		}
	}
	if (edit.field !== 'direction') {
		const parsed = new PptxNativeAnimationService().parseNativeAnimations({
			'p:sld': { 'p:timing': timing },
		});
		const byId = new Map(parsed?.map((item) => [item.nodeId, item]));
		nativeAnimations = nativeAnimations.map((item) => {
			const current = byId.get(item.nodeId);
			if (
				!current ||
				(current.parGroupIndex === item.parGroupIndex &&
					current.parGroupDelayMs === item.parGroupDelayMs &&
					current.groupAutoStart === item.groupAutoStart)
			)
				return item;
			return {
				...item,
				parGroupIndex: current.parGroupIndex,
				parGroupDelayMs: current.parGroupDelayMs,
				groupAutoStart: current.groupAutoStart,
			};
		});
	}
	return { rawTiming: timing, nativeAnimations };
}
