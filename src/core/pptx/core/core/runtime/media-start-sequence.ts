import type { MediaPptxElement, XmlObject } from '../../types';
import {
	appendContainer,
	countTimeNodeChildren,
	ensureMainSequence,
	indexEffectNodes,
	removeEffectNode,
} from '../../services/animation-timing-tree';
import type { EffectNodeRef } from '../../services/animation-timing-tree';
import { ensureArray } from '../../services/native-animation-helpers';
import {
	buildClickGroupNode,
	buildMainSequenceNode,
} from '../../services/animation-write-sequence-builders';
import { xmlChild } from '../../utils/xml-access';

type StartMode = NonNullable<MediaPptxElement['startMode']>;

function mediaPlayEffects(timing: XmlObject): Map<string, EffectNodeRef[]> {
	const effects = new Map<string, EffectNodeRef[]>();
	for (const effect of indexEffectNodes(timing)) {
		if (
			!effect.spid ||
			effect.presetClass !== 'mediacall' ||
			!effect.chain.some((link) => xmlChild(link.node, 'p:cTn')?.['@_nodeType'] === 'mainSeq')
		)
			continue;
		const children = xmlChild(effect.cTn, 'p:childTnLst');
		const command = xmlChild(children, 'p:cmd');
		if (
			!children ||
			!command ||
			countTimeNodeChildren(children) !== 1 ||
			!/^play(?:From\([^)]*\))?$/iu.test(String(command['@_cmd']))
		)
			continue;
		const matched = effects.get(effect.spid) ?? [];
		matched.push(effect);
		effects.set(effect.spid, matched);
	}
	return effects;
}

function startOf(effect: EffectNodeRef): StartMode {
	return effect.cTn['@_nodeType'] === 'withEffect' ? 'auto' : 'sequence';
}

export function readMediaSequenceStarts(timing: XmlObject): Map<string, StartMode> {
	const starts = new Map<string, StartMode>();
	for (const [id, effects] of mediaPlayEffects(timing)) {
		const first = effects[0];
		if (first) starts.set(id, startOf(first));
	}
	return starts;
}

/** Edit the clip's main-sequence Play command and keep its interactive and other effects. */
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
		if (existing.some((effect) => startOf(effect) === start)) continue;
		existing.forEach(removeEffectNode);
		if (start === 'click') continue;

		const main = ensureMainSequence(timing, allocateId, (id) => buildMainSequenceNode(id, []));
		if (!main) continue;
		const duration =
			(media.metadata?.duration ?? 0) * 1000 - (media.trimStartMs ?? 0) - (media.trimEndMs ?? 0);
		const allocateTimeNode = (attributes: XmlObject): XmlObject => ({
			'@_id': String(allocateId()),
			'@_fill': 'hold',
			...attributes,
		});
		const immediate = { 'p:cond': { '@_delay': '0' } };
		const effect: XmlObject = {
			'p:cTn': allocateTimeNode({
				'@_presetID': '1',
				'@_presetClass': 'mediacall',
				'@_presetSubtype': '0',
				'@_nodeType': start === 'auto' ? 'withEffect' : 'clickEffect',
				'p:stCondLst': immediate,
				'p:childTnLst': {
					'p:cmd': {
						'@_type': 'call',
						'@_cmd': 'playFrom(0.0)',
						'p:cBhvr': {
							'p:cTn': allocateTimeNode({
								'@_dur':
									Number.isFinite(duration) && duration > 0 ? String(Math.round(duration)) : '1',
							}),
							'p:tgtEl': { 'p:spTgt': { '@_spid': id } },
						},
					},
				},
			}),
		};
		const parallel: XmlObject = {
			'p:cTn': allocateTimeNode({
				'p:stCondLst': immediate,
				'p:childTnLst': { 'p:par': effect },
			}),
		};
		const group = buildClickGroupNode([parallel], allocateId);
		const children = xmlChild(main, 'p:childTnLst') ?? {};
		main['p:childTnLst'] = children;
		if (start === 'auto') {
			const cTn = xmlChild(group, 'p:cTn')!;
			cTn['p:stCondLst'] = immediate;
			const previous = ensureArray(children['p:par']);
			children['p:par'] = previous.length ? [group, ...previous] : group;
		} else {
			appendContainer(children, 'p:par', group);
		}
	}
}
