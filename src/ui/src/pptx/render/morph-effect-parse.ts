/**
 * Reading the effect CSS a morph carries: `box-shadow` layers and the
 * `drop-shadow()` functions of a `filter`, into layers it can scale and fade.
 *
 * @module render/morph-effect-parse
 */
import { splitCssList, splitCssTokens } from './css-top-level-split';
import type { ShadowLayer } from './morph-effect-shadow';

const LENGTH = /^-?(?:\d+\.?\d*|\.\d+)(?:px)?$/u;

/** Parse a CSS colour this module can fade: `rgb()`, `rgba()` or hex. */
function parseColor(token: string): [number, number, number, number] | undefined {
	const fn = /^rgba?\(([^)]*)\)$/u.exec(token);
	if (fn) {
		const parts = (fn[1] ?? '')
			.split(/[\s,/]+/u)
			.filter(Boolean)
			.map(Number);
		if (parts.length < 3 || parts.some((part) => !Number.isFinite(part))) return undefined;
		return [parts[0] ?? 0, parts[1] ?? 0, parts[2] ?? 0, parts[3] ?? 1];
	}
	const hex = /^#([0-9a-f]{3,8})$/iu.exec(token)?.[1];
	if (!hex || ![3, 4, 6, 8].includes(hex.length)) return undefined;
	const full = hex.length <= 4 ? [...hex].map((c) => c + c).join('') : hex;
	const at = (i: number) => Number.parseInt(full.slice(i, i + 2), 16);
	return [at(0), at(2), at(4), full.length === 8 ? at(6) / 255 : 1];
}

/** `box-shadow` layers, `[]` for none, or `undefined` when a layer cannot be read. */
export function parseShadows(value: string | undefined): ShadowLayer[] | undefined {
	if (!value || value === 'none') return [];
	const layers: ShadowLayer[] = [];
	for (const layer of splitCssList(value)) {
		let inset = false;
		const lengths: number[] = [];
		let color: ShadowLayer['color'] | undefined;
		for (const token of splitCssTokens(layer.trim())) {
			if (token === 'inset') inset = true;
			else if (LENGTH.test(token)) lengths.push(Number.parseFloat(token));
			else color = parseColor(token);
		}
		if (!color || lengths.length < 2) return undefined;
		layers.push({ inset, lengths, color });
	}
	return layers;
}

/** A layer's lengths scaled into a frame `kx` x `ky` times the one it was written in. */
export function scaleLayer(layer: ShadowLayer, kx: number, ky: number): ShadowLayer {
	const mean = Math.sqrt(kx * ky);
	return {
		...layer,
		lengths: layer.lengths.map((value, index) =>
			index === 0 ? value * kx : index === 1 ? value * ky : value * mean,
		),
	};
}

/** The same layer in another colour alpha. */
export function withAlpha(layer: ShadowLayer, alpha: number): ShadowLayer {
	const [r, g, b] = layer.color;
	return { ...layer, color: [r, g, b, alpha] };
}

/** `drop-shadow()` functions of a filter, or `undefined` when it has anything else. */
export function parseDropShadows(value: string | undefined): ShadowLayer[] | undefined {
	if (!value || value === 'none') return [];
	const layers: ShadowLayer[] = [];
	for (const fn of value.match(/[a-z-]+\([^()]*(?:\([^()]*\)[^()]*)*\)/gu) ?? []) {
		const inner = /^drop-shadow\((.*)\)$/u.exec(fn)?.[1];
		const parsed = inner === undefined ? undefined : parseShadows(inner);
		if (!parsed || parsed.length !== 1 || !parsed[0]) return undefined;
		layers.push(parsed[0]);
	}
	return layers;
}
