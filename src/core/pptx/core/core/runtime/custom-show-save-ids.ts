import type { PptxSlide } from '../../types';
import { mapCustomShowActions, normalizeCustomShowIds } from '../../utils/custom-show-actions';
import type { PptxHandlerSaveOptions } from '../types';

/**
 * `p:custShow/@id` is an `xsd:unsignedInt`. When the caller's custom shows carry any other id
 * (a UUID from an editor, say), give those shows the next free number and carry the change to
 * everything that names a show: `p:showPr/p:custShow/@id` and `ppaction://customshow?id=`
 * actions and hyperlinks on the slides. Nothing changes when every id is already valid.
 */
export function withValidCustomShowIds(
	slides: PptxSlide[],
	options: PptxHandlerSaveOptions | undefined,
): { slides: PptxSlide[]; options: PptxHandlerSaveOptions | undefined } {
	if (!options?.customShows) {
		return { slides, options };
	}
	const { shows, idMap } = normalizeCustomShowIds(options.customShows);
	if (idMap.size === 0) {
		return { slides, options };
	}
	const properties = options.presentationProperties;
	const showId = properties?.showSlidesCustomShowId;
	return {
		slides: mapCustomShowActions(slides, (id) => idMap.get(id) ?? id).slides,
		options: {
			...options,
			customShows: shows,
			...(properties && showId !== undefined && idMap.has(showId)
				? { presentationProperties: { ...properties, showSlidesCustomShowId: idMap.get(showId) } }
				: {}),
		},
	};
}
