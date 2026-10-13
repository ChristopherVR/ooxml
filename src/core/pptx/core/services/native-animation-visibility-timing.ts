import type { PptxNativeAnimation, PptxVisibilitySetTiming, XmlObject } from '../types';
import { extractBehaviors } from './native-animation-behaviors';
import { assertAfterAnimationClearSupported } from './animation-after-effect-write';
import { extractAfterAnimationFromSubTnLst } from './native-animation-after-effect';
import { ensureArray, isXmlObject } from './native-animation-helpers';

const neutralTiming = (node: XmlObject) =>
	[
		'@_repeatDur',
		'@_tmFilter',
		'@_endSync',
		'@_evtFilter',
		'@_display',
		'@_masterRel',
		'@_syncBehavior',
		'@_nodePh',
		'@_progress',
		'@_p14:presetBounceEnd',
		'p:endCondLst',
		'p:endSync',
		'p:iterate',
	].every((key) => node[key] === undefined) &&
	(node['@_repeatCount'] === undefined || String(node['@_repeatCount']) === '1000') &&
	(node['@_autoRev'] === undefined || ['0', 'false'].includes(String(node['@_autoRev']))) &&
	(node['@_spd'] === undefined || String(node['@_spd']) === '100000') &&
	(node['@_accel'] === undefined || String(node['@_accel']) === '0') &&
	(node['@_decel'] === undefined || String(node['@_decel']) === '0');

/** Recognize retained visibility bodies even when their value variant cannot be decoded. */
export function hasNativeVisibilitySet(node: XmlObject): boolean {
	if (node['@_presetClass'] !== 'entr' || String(node['@_presetID']) !== '0') return false;
	const children = node['p:childTnLst'];
	if (!isXmlObject(children)) return false;
	return ensureArray(children['p:set']).some((set) => {
		const common = set['p:cBhvr'];
		const list = isXmlObject(common) ? common['p:attrNameLst'] : undefined;
		const names = isXmlObject(list) ? list['p:attrName'] : undefined;
		return (Array.isArray(names) ? names : [names]).some(
			(name) =>
				String(isXmlObject(name) ? name['#text'] : name)
					.trim()
					.toLowerCase() === 'style.visibility',
		);
	});
}

function simpleStarts(node: XmlObject, clickGate = false): boolean {
	const list = node['p:stCondLst'];
	if (list === undefined) return true;
	if (!isXmlObject(list) || Object.keys(list).some((key) => key !== 'p:cond')) return false;
	const conditions = ensureArray(list['p:cond']);
	return (
		conditions.length === 1 &&
		conditions.every(
			(condition) =>
				Object.keys(condition).every((key) => key === '@_delay') &&
				(condition['@_delay'] === undefined ||
					(/^\d+$/u.test(String(condition['@_delay'])) &&
						Number.isSafeInteger(Number(condition['@_delay'])) &&
						Number(condition['@_delay']) <= 2147483647) ||
					(clickGate && condition['@_delay'] === 'indefinite')),
		)
	);
}

/** Conservatively refuse unmeasured enclosing clocks without changing other effects. */
export function unsupportedVisibilityContainerTiming(node: XmlObject): boolean {
	if (node['@_presetClass'] !== undefined) return false;
	const structuralRoot = ['tmRoot', 'mainSeq'].includes(String(node['@_nodeType'] ?? ''));
	return (
		!neutralTiming(node) ||
		!simpleStarts(node, true) ||
		ensureArray(isXmlObject(node['p:stCondLst']) ? node['p:stCondLst']['p:cond'] : undefined).some(
			(condition) =>
				/^\d+$/u.test(String(condition['@_delay'])) && Number(condition['@_delay']) > 0,
		) ||
		(node['@_dur'] !== undefined && !(structuralRoot && node['@_dur'] === 'indefinite')) ||
		(node['@_fill'] !== undefined && node['@_fill'] !== 'hold') ||
		node['p:subTnLst'] !== undefined
	);
}

/** Native preset-zero entrance: one held visibility flip, not a fade approximation. */
export function extractVisibilitySetTiming(node: XmlObject): PptxVisibilitySetTiming | undefined {
	if (
		node['@_presetClass'] !== 'entr' ||
		String(node['@_presetID']) !== '0' ||
		String(node['@_presetSubtype'] ?? '0') !== '0' ||
		node['@_dur'] !== undefined ||
		node['@_fill'] !== 'hold' ||
		!neutralTiming(node) ||
		!simpleStarts(node)
	)
		return undefined;
	const children = node['p:childTnLst'];
	if (!isXmlObject(children) || Object.keys(children).length !== 1 || !children['p:set'])
		return undefined;
	const sets = ensureArray(children['p:set']);
	if (sets.length !== 1) return undefined;
	if (Object.keys(sets[0]!).some((key) => !['p:cBhvr', 'p:to'].includes(key))) return undefined;
	const to = sets[0]!['p:to'];
	if (
		!isXmlObject(to) ||
		Object.keys(to).length !== 1 ||
		!isXmlObject(to['p:strVal']) ||
		Object.keys(to['p:strVal']).some((key) => key !== '@_val') ||
		to['p:strVal']['@_val'] !== 'visible'
	)
		return undefined;
	const common = sets[0]!['p:cBhvr'];
	if (
		!isXmlObject(common) ||
		Object.keys(common).some(
			(key) => !['p:cTn', 'p:tgtEl', 'p:attrNameLst', '@_additive'].includes(key),
		) ||
		(common['@_additive'] !== undefined &&
			!['base', 'repl', 'none'].includes(String(common['@_additive'])))
	)
		return undefined;
	const target = common['p:tgtEl'];
	if (
		!isXmlObject(target) ||
		Object.keys(target).length !== 1 ||
		!isXmlObject(target['p:spTgt']) ||
		Object.keys(target['p:spTgt']).some((key) => key !== '@_spid') ||
		!/^\d+$/u.test(String(target['p:spTgt']['@_spid']))
	)
		return undefined;
	const child = common['p:cTn'];
	if (
		!isXmlObject(child) ||
		!neutralTiming(child) ||
		!simpleStarts(child) ||
		String(child['@_dur']) !== '1' ||
		child['@_fill'] !== 'hold'
	)
		return undefined;
	const behavior = extractBehaviors(children)?.[0];
	if (
		!behavior ||
		behavior.kind !== 'set' ||
		behavior.value !== 'visible' ||
		behavior.attrNames.length !== 1 ||
		behavior.attrNames[0] !== 'style.visibility' ||
		behavior.timing.unsupportedSetTiming
	)
		return undefined;
	try {
		assertAfterAnimationClearSupported(node);
	} catch {
		return undefined;
	}
	const sub = node['p:subTnLst'];
	if (sub !== undefined) {
		if (!isXmlObject(sub)) return undefined;
		for (const [key, raw] of Object.entries(sub)) {
			for (const value of ensureArray(raw)) {
				if (!extractAfterAnimationFromSubTnLst({ 'p:subTnLst': { [key]: value } as XmlObject }))
					return undefined;
			}
		}
	}
	const delayMs = behavior.timing.delayMs ?? 0;
	// The shared playback clock uses finite signed timer intervals.
	if (!Number.isSafeInteger(delayMs) || delayMs < 0 || delayMs >= 2147483647) return undefined;
	return { delayMs, durationMs: 1 };
}

/** Validate the public model before a renderer consumes its admitted timing metadata. */
export function resolveVisibilitySetTiming(
	anim: PptxNativeAnimation,
): PptxVisibilitySetTiming | undefined {
	const timing = anim.visibilitySetTiming;
	if (
		!timing ||
		anim.presetClass !== 'entr' ||
		anim.presetId !== 0 ||
		(anim.presetSubtype ?? 0) !== 0 ||
		anim.fill !== 'hold' ||
		anim.seqConcurrent !== true ||
		anim.seqNextAction !== 'seek' ||
		!['onClick', 'withPrevious'].includes(anim.trigger ?? '') ||
		anim.textTarget ||
		!anim.target ||
		anim.target.type !== 'shape' ||
		anim.target.backgroundOnly ||
		anim.target.subShapeId ||
		anim.target.graphicElement ||
		anim.target.oleChartElement ||
		anim.autoReverse ||
		(anim.accel ?? 0) !== 0 ||
		(anim.decel ?? 0) !== 0 ||
		(anim.startConditions?.length ?? 0) > 1 ||
		anim.startConditions?.some(
			(condition) =>
				Object.entries(condition).some(([key, value]) => value !== undefined && key !== 'delay') ||
				(condition.delay !== undefined &&
					(!Number.isSafeInteger(condition.delay) || condition.delay < 0)),
		) ||
		(anim.repeatCount ?? 1) !== 1 ||
		anim.repeatDurMs !== undefined ||
		(anim.speedPct ?? 100) !== 100 ||
		anim.unsupportedOpacitySetTiming ||
		anim.unsupportedVisibilitySet ||
		anim.endConditions?.length ||
		anim.rawEndCondLst !== undefined ||
		anim.cTnAttributes?.['@_tmFilter'] !== undefined ||
		anim.cTnAttributes?.['@_endSync'] !== undefined ||
		Object.keys(anim.cTnAttributes ?? {}).some((key) => key !== 'p:subTnLst') ||
		!Number.isSafeInteger(timing.delayMs) ||
		timing.delayMs < 0 ||
		timing.delayMs >= 2147483647 ||
		timing.durationMs !== 1 ||
		anim.durationMs !== timing.delayMs + 1
	)
		return undefined;
	const behavior = anim.behaviors?.[0];
	if (
		anim.behaviors?.length !== 1 ||
		!behavior ||
		behavior.kind !== 'set' ||
		behavior.value !== 'visible' ||
		behavior.attrNames.length !== 1 ||
		behavior.attrNames[0] !== 'style.visibility' ||
		behavior.timing.unsupportedSetTiming ||
		behavior.timing.autoReverse ||
		(behavior.timing.repeatCount ?? 1) !== 1 ||
		behavior.timing.tmFilter !== undefined ||
		(behavior.timing.accel ?? 0) !== 0 ||
		(behavior.timing.decel ?? 0) !== 0 ||
		(behavior.additive !== undefined && !['base', 'repl', 'none'].includes(behavior.additive)) ||
		behavior.timing.fill !== 'hold' ||
		behavior.timing.durationMs !== 1 ||
		(behavior.timing.delayMs ?? 0) !== timing.delayMs
	)
		return undefined;
	return timing;
}
