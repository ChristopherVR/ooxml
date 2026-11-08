/**
 * Map the SDK's picture tile and stretch inputs (`ImageOptions.tile` /
 * `fillRect`, `FillInput { type: 'image' }`) onto the picture layout fields
 * of the model ({@link PptxImageLayout}), refusing values the file cannot
 * hold or the loader would not read back.
 *
 * @module sdk/picture-layout-input
 */

import type { PptxImageLayout } from '../../types/image';
import type { PictureTileInput, RelativeRectInput } from './shape-style-input-types';

/** What a picture or a picture fill says about its layout. */
export interface PictureLayoutInput {
	mode?: 'stretch' | 'tile' | undefined;
	tile?: PictureTileInput | undefined;
	fillRect?: RelativeRectInput | undefined;
}

const FLIPS: ReadonlySet<string> = new Set(['none', 'x', 'y', 'xy']);
const ALIGNMENTS: ReadonlySet<string> = new Set([
	'tl',
	't',
	'tr',
	'l',
	'ctr',
	'r',
	'bl',
	'b',
	'br',
]);
/** The loader caps `a:fillRect` insets at ten frames (`parseBoundedRectFraction`). */
const MAX_INSET = 10;

function finite(where: string, name: string, value: unknown): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw new RangeError(`${where}.${name} must be a finite number, got ${String(value)}`);
	}
	return value;
}

function tileLayout(where: string, tile: PictureTileInput): PptxImageLayout {
	const scale = (name: 'scaleX' | 'scaleY'): number => {
		const value = finite(where, `tile.${name}`, tile[name] ?? 1);
		if (value <= 0) {
			throw new RangeError(`${where}.tile.${name} must be above 0, got ${value}`);
		}
		return value;
	};
	const flip = tile.flip ?? 'none';
	if (!FLIPS.has(flip)) {
		throw new RangeError(`${where}.tile.flip must be none, x, y or xy, got ${String(flip)}`);
	}
	const alignment = tile.alignment ?? 'tl';
	if (!ALIGNMENTS.has(alignment)) {
		throw new RangeError(
			`${where}.tile.alignment must be one of ${[...ALIGNMENTS].join(', ')}, got ${String(alignment)}`,
		);
	}
	return {
		tileOffsetX: finite(where, 'tile.offsetX', tile.offsetX ?? 0),
		tileOffsetY: finite(where, 'tile.offsetY', tile.offsetY ?? 0),
		tileScaleX: scale('scaleX'),
		tileScaleY: scale('scaleY'),
		tileFlip: flip,
		tileAlignment: alignment,
	};
}

function fillRectLayout(where: string, rect: RelativeRectInput): PptxImageLayout {
	const inset = (name: 'l' | 't' | 'r' | 'b'): number => {
		const value = finite(where, `fillRect.${name}`, rect[name]);
		if (Math.abs(value) > MAX_INSET) {
			throw new RangeError(
				`${where}.fillRect.${name} must be between -${MAX_INSET} and ${MAX_INSET}, got ${value}`,
			);
		}
		return value;
	};
	const layout = {
		fillRectLeft: inset('l'),
		fillRectTop: inset('t'),
		fillRectRight: inset('r'),
		fillRectBottom: inset('b'),
	};
	if (layout.fillRectLeft + layout.fillRectRight >= 1) {
		throw new RangeError(`${where}.fillRect.l + r must be below 1, which leaves no picture`);
	}
	if (layout.fillRectTop + layout.fillRectBottom >= 1) {
		throw new RangeError(`${where}.fillRect.t + b must be below 1, which leaves no picture`);
	}
	return layout;
}

/**
 * The layout fields for `input`, or `undefined` for a plain stretch (the
 * writer's default, an empty `a:fillRect`). A tile gets every field, with
 * PowerPoint's defaults (no offset, 100% scale, no flip, top-left).
 *
 * @param where - The option path named in errors (e.g. `ImageOptions`).
 * @throws RangeError for a tile and a fill rectangle together, either one
 *   against the `mode`, an unknown `mode`, or a value outside its range.
 */
export function mapPictureLayoutInput(
	where: string,
	input: PictureLayoutInput,
): PptxImageLayout | undefined {
	const { mode, tile, fillRect } = input;
	if (mode !== undefined && mode !== 'stretch' && mode !== 'tile') {
		throw new RangeError(`${where}.mode must be 'stretch' or 'tile', got ${String(mode)}`);
	}
	if (tile && fillRect) {
		throw new RangeError(`${where} takes a tile or a fillRect, not both`);
	}
	if (tile && mode === 'stretch') {
		throw new RangeError(`${where}.tile needs mode 'tile', got 'stretch'`);
	}
	if (fillRect && mode === 'tile') {
		throw new RangeError(`${where}.fillRect needs mode 'stretch', got 'tile'`);
	}
	if (tile || mode === 'tile') {
		return tileLayout(where, tile ?? {});
	}
	return fillRect ? fillRectLayout(where, fillRect) : undefined;
}
