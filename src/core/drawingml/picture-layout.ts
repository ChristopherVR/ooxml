import { NS, first, type XmlElement } from '../xml/index';
import { picturePercentage } from './picture-effects';
export interface PictureInsets {
	l: number;
	t: number;
	r: number;
	b: number;
}
export interface PictureTile {
	scaleX?: number;
	scaleY?: number;
	offsetXEmu?: number;
	offsetYEmu?: number;
	alignment?: string;
	flip?: 'none' | 'x' | 'y' | 'xy';
}
export interface DrawingPictureLayout {
	tile?: PictureTile;
	fillRect?: PictureInsets;
	rotWithShape?: boolean;
	dpi?: number;
}
export function parsePictureLayout(fill: XmlElement): DrawingPictureLayout {
	const out: DrawingPictureLayout = {};
	const rot = fill.getAttribute('rotWithShape');
	if (rot !== null) out.rotWithShape = rot !== '0' && rot !== 'false';
	const dpi = Number(fill.getAttribute('dpi'));
	if (fill.hasAttribute('dpi') && Number.isFinite(dpi) && dpi > 0) out.dpi = dpi;
	const tile = first(fill, 'tile', NS.a);
	if (tile) {
		const value: PictureTile = {};
		for (const [key, attr] of [
			['scaleX', 'sx'],
			['scaleY', 'sy'],
		] as const) {
			const n = picturePercentage(tile.getAttribute(attr));
			if (n !== undefined) value[key] = n;
		}
		for (const [key, attr] of [
			['offsetXEmu', 'tx'],
			['offsetYEmu', 'ty'],
		] as const) {
			const raw = tile.getAttribute(attr);
			const n = raw === null ? NaN : Number(raw);
			if (Number.isSafeInteger(n)) value[key] = n;
		}
		const alignment = tile.getAttribute('algn');
		if (alignment) value.alignment = alignment;
		const flip = tile.getAttribute('flip');
		if (flip && ['none', 'x', 'y', 'xy'].includes(flip))
			value.flip = flip as NonNullable<PictureTile['flip']>;
		out.tile = value;
	}
	const rect = first(first(fill, 'stretch', NS.a), 'fillRect', NS.a);
	if (rect)
		out.fillRect = {
			l: picturePercentage(rect.getAttribute('l')) ?? 0,
			t: picturePercentage(rect.getAttribute('t')) ?? 0,
			r: picturePercentage(rect.getAttribute('r')) ?? 0,
			b: picturePercentage(rect.getAttribute('b')) ?? 0,
		};
	return out;
}
/** Updates owned placement fields only; unknown attributes/extensions stay source-backed. */
export function applyPictureLayout(fill: XmlElement, layout: DrawingPictureLayout): void {
	const before = parsePictureLayout(fill);
	const attr = (
		node: XmlElement,
		name: string,
		value: string | number | undefined,
		old: unknown,
	) => {
		if (value === old) return;
		if (value === undefined) node.removeAttribute(name);
		else node.setAttribute(name, String(value));
	};
	if (layout.rotWithShape !== before.rotWithShape)
		attr(
			fill,
			'rotWithShape',
			layout.rotWithShape === undefined ? undefined : layout.rotWithShape ? '1' : '0',
			undefined,
		);
	if (layout.dpi !== before.dpi) {
		if (layout.dpi !== undefined && (!Number.isSafeInteger(layout.dpi) || layout.dpi <= 0))
			throw new RangeError('Invalid picture DPI');
		attr(fill, 'dpi', layout.dpi, undefined);
	}
	const make = (parent: XmlElement, name: string) => {
		const old = first(parent, name, NS.a);
		if (old) return old;
		const node = fill.ownerDocument.createElementNS(NS.a, `a:${name}`);
		parent.insertBefore(node, first(parent, 'extLst', NS.a) ?? null);
		return node;
	};
	if (layout.tile) {
		const stretch = first(fill, 'stretch', NS.a);
		if (stretch) fill.removeChild(stretch);
		const tile = make(fill, 'tile');
		for (const [key, name] of [
			['scaleX', 'sx'],
			['scaleY', 'sy'],
			['offsetXEmu', 'tx'],
			['offsetYEmu', 'ty'],
		] as const) {
			const n = layout.tile[key];
			if (
				n !== undefined &&
				(!Number.isFinite(n) ||
					(name === 'sx' || name === 'sy' ? n <= 0 : !Number.isSafeInteger(n)))
			)
				throw new RangeError('Invalid picture tile');
			if (n !== before.tile?.[key])
				attr(
					tile,
					name,
					n === undefined
						? undefined
						: Math.round(n * (name === 'sx' || name === 'sy' ? 100000 : 1)),
					undefined,
				);
		}
		if (layout.tile.alignment !== before.tile?.alignment) {
			if (
				layout.tile.alignment &&
				!['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(layout.tile.alignment)
			)
				throw new RangeError('Invalid picture tile alignment');
			attr(tile, 'algn', layout.tile.alignment, undefined);
		}
		if (layout.tile.flip !== undefined && !['none', 'x', 'y', 'xy'].includes(layout.tile.flip))
			throw new RangeError('Invalid picture tile flip');
		if (layout.tile.flip !== before.tile?.flip) attr(tile, 'flip', layout.tile.flip, undefined);
	} else if (before.tile || layout.fillRect) {
		const tile = first(fill, 'tile', NS.a);
		if (tile) fill.removeChild(tile);
		const rect = make(make(fill, 'stretch'), 'fillRect');
		for (const key of ['l', 't', 'r', 'b'] as const) {
			const n = layout.fillRect?.[key];
			if (n !== undefined && !Number.isFinite(n))
				throw new RangeError('Invalid picture fill rectangle');
			if (n !== before.fillRect?.[key])
				attr(rect, key, n === undefined ? undefined : Math.round(n * 100000), undefined);
		}
	}
}
