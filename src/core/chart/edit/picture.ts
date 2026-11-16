// Picture fills of a Format pane target (Picture or texture fill): a picture chosen as a `data:`
// URL, its transparency and whether it is tiled. A new picture has no relationship yet, so it is
// filed under a pending id in `ChartSpace.pictureSources`, which the painter reads at once and
// the host's save replaces with a real relationship (`chart/picture-parts`: `embedChartSpacePictures`).
import type { DrawingFill, DrawingPictureEffect } from '../../drawingml/index';
import type { ChartSpace } from '../model';
import { chartTargetShape, setChartTargetFill } from './target-paint';
import type { ChartFormatTarget } from './target';

/** What a picture fill edit sets; omitted fields keep the target's current picture settings. */
export interface ChartPictureEdit {
	/** A new picture, as a `data:image/...` URL. */
	src?: string;
	/** Tile the picture as a texture (`a:tile`) instead of stretching it (`a:stretch`). */
	tile?: boolean;
	/** Transparency 0..100 (`a:alphaModFix`). */
	transparency?: number;
}

/** The prefix of the ids a pending picture is filed under until the host saves it. */
export const CHART_PENDING_PICTURE_PREFIX = 'chartPicturePending';

/** Whether `relId` names a picture the host has not stored in its package yet. */
export const isPendingChartPicture = (relId: string | undefined): boolean =>
	!!relId?.startsWith(CHART_PENDING_PICTURE_PREFIX);

const DATA_URL =
	/^data:image\/(?:png|jpe?g|gif|bmp|tiff|webp|svg\+xml);base64,[A-Za-z0-9+/=\s]+$/iu;

/** The id `src` is filed under: one it already has, else the next free pending id. */
function pictureId(sources: Record<string, string>, src: string): string {
	const known = Object.entries(sources).find(([, value]) => value === src);
	if (known) return known[0];
	let n = 1;
	while (sources[`${CHART_PENDING_PICTURE_PREFIX}${n}`] !== undefined) n++;
	return `${CHART_PENDING_PICTURE_PREFIX}${n}`;
}

/**
 * Gives `target` a picture fill. With `src` the picture is replaced; without it the target must
 * already have a picture fill, whose tiling and transparency change. Throws a `RangeError` for a
 * URL that is not a base64 picture, a transparency outside 0..100, or nothing to change.
 */
export function setChartTargetPicture(
	space: ChartSpace,
	target: ChartFormatTarget,
	edit: ChartPictureEdit,
): ChartSpace {
	const current = chartTargetShape(space, target)?.fill;
	const before = current?.kind === 'picture' ? current : undefined;
	if (edit.src !== undefined && !DATA_URL.test(edit.src))
		throw new RangeError('A chart picture must be a base64 image data URL');
	if (
		edit.transparency !== undefined &&
		(!Number.isFinite(edit.transparency) || edit.transparency < 0 || edit.transparency > 100)
	)
		throw new RangeError('Invalid picture transparency');
	const sources = { ...space.pictureSources };
	let relId = before?.relId;
	if (edit.src !== undefined) {
		relId = pictureId(sources, edit.src);
		sources[relId] = edit.src;
	}
	if (!relId) throw new RangeError('The target has no picture to change');
	const fill: Extract<DrawingFill, { kind: 'picture' }> = {
		...(before && edit.src === undefined ? before : {}),
		kind: 'picture',
		relId,
	};
	const tiled = edit.tile ?? Boolean(before?.layout?.tile);
	const layout = { ...before?.layout };
	if (tiled) layout.tile = { ...before?.layout?.tile };
	else {
		delete layout.tile;
		delete layout.fillRect;
	}
	fill.layout = layout;
	const alpha = before?.effects?.find((effect) => effect.kind === 'alphaModFix');
	const transparency =
		edit.transparency ?? (alpha?.kind === 'alphaModFix' ? Math.round((1 - alpha.amount) * 100) : 0);
	const effects: DrawingPictureEffect[] = (before?.effects ?? []).filter(
		(effect) => effect.kind !== 'alphaModFix',
	);
	if (transparency > 0) effects.push({ kind: 'alphaModFix', amount: (100 - transparency) / 100 });
	fill.effects = effects;
	if (edit.src !== undefined) delete fill.sourceXml;
	const next = setChartTargetFill(space, target, fill);
	return { ...next, pictureSources: sources };
}
