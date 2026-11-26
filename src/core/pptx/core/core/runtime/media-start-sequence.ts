import type { MediaPptxElement, XmlObject } from '../../types';
import {
	appendContainer,
	countTimeNodeChildren,
	ensureMainSequence,
	indexEffectNodes,
	removeEffectNode,
} from '../../services/animation-timing-tree';
import type { EffectNodeRef } from '../../services/animation-timing-tree';
import { parseTimeTargetElement } from '../../services/animation-target-build-helpers';
import { ensureArray } from '../../services/native-animation-helpers';
import {
	buildClickGroupNode,
	buildMainSequenceNode,
	wrapInInteractiveSequence,
} from '../../services/animation-write-sequence-builders';
import { xmlChild } from '../../utils/xml-access';

type StartMode = NonNullable<MediaPptxElement['startMode']>;
type MediaEffect = EffectNodeRef & { mode: StartMode; toggle: boolean };

function conditions(node: XmlObject | undefined): XmlObject[] {
	return ensureArray(xmlChild(node, 'p:stCondLst')?.['p:cond']);
}

function mediaPlayEffects(timing: XmlObject): Map<string, MediaEffect[]> {
	const effects = new Map<string, MediaEffect[]>();
	const clickGroups = new Set<XmlObject>();
	for (const effect of indexEffectNodes(timing)) {
		const sequenceIndex = effect.chain.findIndex(
			({ key, node }) =>
				key === 'p:seq' &&
				['mainSeq', 'interactiveSeq'].includes(String(xmlChild(node, 'p:cTn')?.['@_nodeType'])),
		);
		const sequence = xmlChild(effect.chain[sequenceIndex]?.node, 'p:cTn');
		const group = effect.chain[sequenceIndex + 1];
		const main = sequence?.['@_nodeType'] === 'mainSeq';
		const type = effect.cTn['@_nodeType'];
		if (main && group && type === 'clickEffect') clickGroups.add(group.node);
		if (!sequence || !effect.spid || effect.presetClass !== 'mediacall') continue;

		const children = xmlChild(effect.cTn, 'p:childTnLst');
		const command = xmlChild(children, 'p:cmd');
		if (!children || !command || countTimeNodeChildren(children) !== 1) continue;
		const cmd = String(command['@_cmd']);
		const play = /^play(?:From\([^)]*\))?$/iu.test(cmd);
		const toggle = /^togglePause$/iu.test(cmd);

		let mode: StartMode;
		if (main && play) {
			const startsAtEntry =
				group &&
				(ensureArray(group.holder[group.key])[0] === group.node ||
					conditions(xmlChild(group.node, 'p:cTn')).some(
						(condition) =>
							condition['@_delay'] === '0' &&
							(Object.keys(condition).length === 1 ||
								(condition['@_evt'] === 'onBegin' &&
									String(xmlChild(condition, 'p:tn')?.['@_val']) === String(sequence['@_id']))),
					));
			mode =
				startsAtEntry &&
				!clickGroups.has(group.node) &&
				(type === 'withEffect' || type === 'afterEffect')
					? 'auto'
					: 'sequence';
		} else if (
			!main &&
			(play || toggle) &&
			conditions(sequence).some((condition) => {
				const target = parseTimeTargetElement(xmlChild(condition, 'p:tgtEl'));
				return (
					condition['@_evt'] === 'onClick' &&
					target?.type === 'shape' &&
					(target.subShapeId ?? target.shapeId) === effect.spid
				);
			})
		) {
			mode = 'click';
		} else {
			continue;
		}
		const matched = effects.get(effect.spid) ?? [];
		matched.push({ ...effect, mode, toggle });
		effects.set(effect.spid, matched);
	}
	return effects;
}

function primaryEffect(effects: MediaEffect[]): MediaEffect | undefined {
	return effects.find(({ mode }) => mode !== 'click') ?? effects[0];
}

export function readMediaSequenceStarts(timing: XmlObject): Map<string, StartMode> {
	const starts = new Map<string, StartMode>();
	for (const [id, effects] of mediaPlayEffects(timing)) {
		const first = primaryEffect(effects);
		if (first) starts.set(id, first.mode);
	}
	return starts;
}

function startCondition(delay: string): XmlObject {
	return { 'p:cond': { '@_delay': delay } };
}

function buildMediaStartGroup(
	media: MediaPptxElement,
	id: string,
	mode: StartMode,
	allocateId: () => number,
): XmlObject {
	const duration =
		(media.metadata?.duration ?? 0) * 1000 - (media.trimStartMs ?? 0) - (media.trimEndMs ?? 0);
	const allocateTimeNode = (attributes: XmlObject): XmlObject => ({
		'@_id': String(allocateId()),
		'@_fill': 'hold',
		...attributes,
	});
	return {
		'p:cTn': allocateTimeNode({
			'p:stCondLst': startCondition('0'),
			'p:childTnLst': {
				'p:par': {
					'p:cTn': allocateTimeNode({
						'@_presetID': mode === 'click' ? '2' : '1',
						'@_presetClass': 'mediacall',
						'@_presetSubtype': '0',
						'@_nodeType': mode === 'auto' ? 'withEffect' : 'clickEffect',
						'p:stCondLst': startCondition('0'),
						'p:childTnLst': {
							'p:cmd': {
								'@_type': 'call',
								'@_cmd': mode === 'click' ? 'togglePause' : 'playFrom(0.0)',
								'p:cBhvr': {
									'p:cTn': allocateTimeNode({
										'@_dur':
											mode !== 'click' && Number.isFinite(duration) && duration > 0
												? String(Math.round(duration))
												: '1',
									}),
									'p:tgtEl': { 'p:spTgt': { '@_spid': id } },
								},
							},
						},
					}),
				},
			},
		}),
	};
}

function appendInteractiveStart(
	timing: XmlObject,
	group: XmlObject,
	id: string,
	allocateId: () => number,
): void {
	const list = xmlChild(timing, 'p:tnLst');
	if (!list) return;
	const root = ensureArray(list['p:par'])
		.map((node) => xmlChild(node, 'p:cTn'))
		.find((node) => node?.['@_nodeType'] === 'tmRoot');
	if (!root) return;
	const seq = wrapInInteractiveSequence([group], id, allocateId);
	seq['@_concurrent'] = '1';
	seq['@_nextAc'] = 'seek';
	const cTn = xmlChild(seq, 'p:cTn')!;
	delete cTn['@_dur'];
	cTn['@_restart'] = 'whenNotActive';
	cTn['@_fill'] = 'hold';
	cTn['@_evtFilter'] = 'cancelBubble';
	const children = cTn['p:childTnLst'];
	delete cTn['p:childTnLst'];
	cTn['p:endSync'] = { '@_evt': 'end', '@_delay': '0', 'p:rtn': { '@_val': 'all' } };
	cTn['p:childTnLst'] = children;
	const holder = xmlChild(root, 'p:childTnLst') ?? {};
	root['p:childTnLst'] = holder;
	appendContainer(holder, 'p:seq', seq);
}

/** Edit the clip's start effect, keeping unrelated timing and the native click-to-pause action. */
export function applyMediaSequenceStarts(
	timing: XmlObject,
	mediaById: ReadonlyMap<string, MediaPptxElement>,
	allocateId: () => number,
): void {
	if (![...mediaById.values()].some((media) => media.startMode !== undefined)) return;
	const effects = mediaPlayEffects(timing);
	for (const [id, media] of mediaById) {
		const start = media.startMode;
		if (!start) continue;
		const existing = effects.get(id) ?? [];
		if (primaryEffect(existing)?.mode === start) continue;
		const hasClick = existing.some(({ mode }) => mode === 'click');
		for (const effect of existing) {
			if (effect.mode !== 'click' || (start !== 'click' && !effect.toggle))
				removeEffectNode(effect);
		}
		if (start === 'click' && hasClick) continue;

		const parallel = buildMediaStartGroup(media, id, start, allocateId);
		if (start === 'click') {
			appendInteractiveStart(timing, parallel, id, allocateId);
			continue;
		}

		const main = ensureMainSequence(timing, allocateId, (id) => buildMainSequenceNode(id, []));
		if (!main) continue;
		const group = buildClickGroupNode([parallel], allocateId);
		const children = xmlChild(main, 'p:childTnLst') ?? {};
		main['p:childTnLst'] = children;
		if (start === 'auto') {
			const cTn = xmlChild(group, 'p:cTn')!;
			cTn['p:stCondLst'] = {
				'p:cond': [
					{ '@_delay': 'indefinite' },
					{ '@_evt': 'onBegin', '@_delay': '0', 'p:tn': { '@_val': main['@_id'] } },
				],
			};
			children['p:par'] = [group, ...ensureArray(children['p:par'])];
		} else {
			appendContainer(children, 'p:par', group);
		}
	}
}
