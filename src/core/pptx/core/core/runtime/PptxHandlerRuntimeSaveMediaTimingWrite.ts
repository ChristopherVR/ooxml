import { XmlObject, PptxElement } from '../../types';
import type { MediaPptxElement } from '../../types';
import { maxTimeNodeId } from '../../services/animation-timing-tree';
import { xmlChild } from '../../utils/xml-access';
import {
	getXmlShapeIdFromXml,
	resolvePlayAcrossSlides,
} from './PptxHandlerRuntimeMediaParsingUtils';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSavePresPropsAndSignatures';
import { applyMediaSequenceStarts } from './media-start-sequence';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	protected collectMediaElements(elements: PptxElement[], output: MediaPptxElement[]): void {
		for (const element of elements) {
			if (element.type === 'media') {
				output.push(element);
			} else if (element.type === 'group' && Array.isArray(element.children)) {
				this.collectMediaElements(element.children, output);
			}
		}
	}

	protected applyMediaTimingToTimingTree(
		node: XmlObject,
		mediaByShapeId: Map<string, MediaPptxElement>,
		seen = new Set<string>(),
		depth = 0,
	): void {
		if (depth > 64) return;
		for (const mediaTag of ['p:video', 'p:audio']) {
			const mediaNodes = this.ensureArray(node[mediaTag]) as XmlObject[];
			mediaNodes.forEach((mediaNode) => {
				const cMediaNode = mediaNode['p:cMediaNode'] as XmlObject | undefined;
				if (!cMediaNode) {
					return;
				}
				const tgtEl = cMediaNode['p:tgtEl'] as XmlObject | undefined;
				const spTgt = tgtEl?.['p:spTgt'] as XmlObject | undefined;
				// `p:subSp/@_spid` names the actual leaf media shape when it sits
				// inside a group (see `PptxAnimationTarget.subShapeId`); `mediaByShapeId`
				// is keyed by each element's OWN id (`collectMediaElements` recurses
				// into group children), so a grouped media shape only matches via the
				// sub-shape id, never the enclosing group's `@_spid`.
				const subSp = spTgt?.['p:subSp'] as XmlObject | undefined;
				const rawShapeId = subSp?.['@_spid'] ?? spTgt?.['@_spid'];
				const shapeId = rawShapeId !== undefined ? String(rawShapeId).trim() : '';
				if (shapeId.length === 0) {
					return;
				}
				const media = mediaByShapeId.get(shapeId);
				if (!media) {
					return;
				}
				seen.add(shapeId);

				let cTn = cMediaNode['p:cTn'] as XmlObject | undefined;
				if (!cTn) {
					cTn = {};
					cMediaNode['p:cTn'] = cTn;
				}

				if (media.loop) {
					cTn['@_repeatCount'] = 'indefinite';
				} else if (String(cTn['@_repeatCount']) === 'indefinite') {
					delete cTn['@_repeatCount'];
				}

				if (media.autoPlay && !media.startMode) {
					delete cTn['@_nodeType'];
					cTn['p:stCondLst'] = { 'p:cond': { '@_delay': '0' } };
				} else if (
					media.startMode !== undefined ||
					media.autoPlay === false ||
					['1', '2'].includes(String(cTn['@_nodeType']))
				) {
					delete cTn['@_nodeType'];
					cTn['p:stCondLst'] = { 'p:cond': { '@_delay': 'indefinite' } };
				}
				if (media.rewindAfterPlaying !== undefined) {
					cTn['@_fill'] = media.rewindAfterPlaying ? 'remove' : 'hold';
				}

				// Play across slides has two storage forms (see
				// `resolvePlayAcrossSlides`): cTn `dur="indefinite"` and PowerPoint's
				// own `cMediaNode/@numSld` span. When the node already says it through
				// `numSld`, adding `dur` too only grows the XML.
				const spansSlides = resolvePlayAcrossSlides(cMediaNode, false, mediaTag);
				if (media.playAcrossSlides && mediaTag === 'p:audio') {
					if (!spansSlides) {
						cMediaNode['@_numSld'] = '999';
					}
				} else if (!media.playAcrossSlides) {
					// Only remove if we're sure it was previously set for play-across
					// Leave dur alone if it was set for other reasons
					if (String(cTn['@_dur']) === 'indefinite') {
						delete cTn['@_dur'];
					}
					// A multi-slide span would read back as "play across" again.
					if (spansSlides) {
						delete cMediaNode['@_numSld'];
					}
				}

				delete cMediaNode['@_fullScrn'];
				if (media.fullScreen && mediaTag === 'p:video') {
					mediaNode['@_fullScrn'] = '1';
				} else {
					delete mediaNode['@_fullScrn'];
				}

				// Volume (0-1 → 0-100000)
				if (media.volume !== undefined && Number.isFinite(media.volume)) {
					cMediaNode['@_vol'] = String(Math.round(Math.min(1, Math.max(0, media.volume)) * 100000));
				}

				// Hide when not playing → showWhenStopped="0"
				if (media.hideWhenNotPlaying) {
					cMediaNode['@_showWhenStopped'] = '0';
				} else if (String(cMediaNode['@_showWhenStopped']) === '0') {
					// Only a stale "0" is cleared; an authored "1" is the default
					// value written explicitly and must survive the rewrite.
					delete cMediaNode['@_showWhenStopped'];
				}
			});
		}

		const cTn = node['p:cTn'] as XmlObject | undefined;
		if (cTn) {
			const childTnLst = cTn['p:childTnLst'] as XmlObject | undefined;
			if (childTnLst) {
				this.applyMediaTimingToTimingTree(childTnLst, mediaByShapeId, seen, depth + 1);
			}
		}

		for (const containerTag of ['p:par', 'p:seq', 'p:excl', 'p:tnLst']) {
			const containerNodes = this.ensureArray(node[containerTag]) as XmlObject[];
			containerNodes.forEach((containerNode) => {
				this.applyMediaTimingToTimingTree(containerNode, mediaByShapeId, seen, depth + 1);
			});
		}
	}

	protected applyMediaTimingToRawTiming(
		rawTiming: XmlObject | undefined,
		elements: PptxElement[],
		ids: ReadonlyMap<string, string>,
	): XmlObject | undefined {
		const mediaElements: MediaPptxElement[] = [];
		this.collectMediaElements(elements, mediaElements);
		if (mediaElements.length === 0) {
			return rawTiming;
		}

		const mediaByShapeId = new Map<string, MediaPptxElement>();
		mediaElements.forEach((mediaElement) => {
			const shapeId = getXmlShapeIdFromXml(mediaElement.rawXml) ?? ids.get(mediaElement.id);
			if (!shapeId) {
				return;
			}
			mediaByShapeId.set(shapeId, mediaElement);
		});
		const seen = new Set<string>();
		const timing = rawTiming ? structuredClone(rawTiming) : {};
		this.applyMediaTimingToTimingTree(timing, mediaByShapeId, seen);
		const missing = [...mediaByShapeId].filter(
			([id, media]) =>
				!seen.has(id) &&
				(media.rawXml?.['p:nvPicPr'] || !media.rawXml) &&
				(media.mediaPath || media.mediaData) &&
				[
					media.autoPlay,
					media.startMode,
					media.volume,
					media.loop,
					media.fullScreen,
					media.hideWhenNotPlaying,
					media.rewindAfterPlaying,
					media.playAcrossSlides,
				].some((value) => value !== undefined),
		);
		let nextId = maxTimeNodeId(timing) + 1;
		const allocateId = () => nextId++;
		if (missing.length === 0) {
			if (rawTiming) applyMediaSequenceStarts(timing, mediaByShapeId, allocateId);
			return rawTiming ? timing : undefined;
		}
		const list = xmlChild(timing, 'p:tnLst') ?? {};
		timing['p:tnLst'] = list;
		const roots = this.ensureArray(list['p:par']) as XmlObject[];
		let root = roots
			.map((node) => xmlChild(node, 'p:cTn'))
			.find((node) => node?.['@_nodeType'] === 'tmRoot');
		if (!root) {
			root = {
				'@_id': String(nextId++),
				'@_dur': 'indefinite',
				'@_restart': 'never',
				'@_nodeType': 'tmRoot',
			};
			list['p:par'] = [...roots, { 'p:cTn': root }];
		}
		const children = xmlChild(root, 'p:childTnLst') ?? {};
		root['p:childTnLst'] = children;
		const added: XmlObject = {};
		for (const [id, media] of missing) {
			const tag = media.mediaType === 'audio' ? 'p:audio' : 'p:video';
			const nodes = this.ensureArray(added[tag]) as XmlObject[];
			added[tag] = [
				...nodes,
				{
					'p:cMediaNode': {
						'p:cTn': {
							'@_id': String(nextId++),
							'@_fill': 'hold',
							'@_display': '0',
							'p:stCondLst': { 'p:cond': { '@_delay': 'indefinite' } },
						},
						'p:tgtEl': { 'p:spTgt': { '@_spid': id } },
					},
				},
			];
		}
		this.applyMediaTimingToTimingTree(added, mediaByShapeId);
		for (const [tag, nodes] of Object.entries(added)) {
			children[tag] = [...this.ensureArray(children[tag]), ...this.ensureArray(nodes)];
		}
		applyMediaSequenceStarts(timing, mediaByShapeId, allocateId);
		return timing;
	}
}
