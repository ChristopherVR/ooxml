/**
 * Computed CSS values read for vector PDF export: colours, transforms, rounded boxes, clip
 * shapes and single linear gradients. Every reader returns `undefined` for a value it does not
 * understand, so the caller can draw that element as a picture instead. Pure.
 */
import { parseColor } from 'ooxml-core/pdf/write';
import type { PdfShading } from 'ooxml-core/pdf/write';
import type { Matrix, Rgba } from './pdf-vector-display-list';

/** A computed colour; `null` for transparent; `undefined` when it cannot be read. */
export function cssColor(value: string | null | undefined): Rgba | null | undefined {
	try {
		const color = parseColor(value);
		return color && color.a > 0 ? color : null;
	} catch {
		return undefined;
	}
}

/**
 * A computed `transform` as a 2-D matrix; `undefined` for a 3-D transform that is not flat
 * (perspective, or depth that changes the picture).
 */
export function cssMatrix(value: string): Matrix | undefined {
	const text = value.trim();
	if (!text || text === 'none') return [1, 0, 0, 1, 0, 0];
	const two = /^matrix\(([^)]*)\)$/.exec(text);
	if (two) {
		const v = two[1]!.split(',').map(Number);
		return v.length === 6 && v.every(Number.isFinite)
			? [v[0]!, v[1]!, v[2]!, v[3]!, v[4]!, v[5]!]
			: undefined;
	}
	const three = /^matrix3d\(([^)]*)\)$/.exec(text);
	if (three) {
		const m = three[1]!.split(',').map(Number);
		if (m.length !== 16 || !m.every(Number.isFinite)) return undefined;
		// Flat only: no perspective row and no z mixing into x or y.
		const flat =
			Math.abs(m[3]!) < 1e-9 &&
			Math.abs(m[7]!) < 1e-9 &&
			Math.abs(m[11]!) < 1e-9 &&
			Math.abs(m[15]! - 1) < 1e-9 &&
			Math.abs(m[8]!) < 1e-9 &&
			Math.abs(m[9]!) < 1e-9;
		return flat ? [m[0]!, m[1]!, m[4]!, m[5]!, m[12]!, m[13]!] : undefined;
	}
	return undefined;
}

/** `transform-origin` in px, relative to the border box. */
export function cssOrigin(value: string, width: number, height: number): [number, number] {
	const parts = value.trim().split(/\s+/);
	const read = (part: string | undefined, size: number) => {
		if (!part) return size / 2;
		if (part.endsWith('%')) return (Number.parseFloat(part) / 100) * size;
		const parsed = Number.parseFloat(part);
		return Number.isFinite(parsed) ? parsed : size / 2;
	};
	return [read(parts[0], width), read(parts[1], height)];
}

/** Rounded rectangle path data; `radii` are [top-left, top-right, bottom-right, bottom-left]. */
export function roundedRectPath(
	x: number,
	y: number,
	w: number,
	h: number,
	radii: readonly [number, number, number, number] = [0, 0, 0, 0],
): string {
	const limit = Math.min(w, h) / 2;
	const [tl, tr, br, bl] = radii.map((r) => Math.max(0, Math.min(r, limit))) as [
		number,
		number,
		number,
		number,
	];
	if (!(tl || tr || br || bl))
		return `M ${x} ${y} L ${x + w} ${y} L ${x + w} ${y + h} L ${x} ${y + h} Z`;
	const k = 0.5522847498;
	return [
		`M ${x + tl} ${y}`,
		`L ${x + w - tr} ${y}`,
		tr ? `C ${x + w - tr + tr * k} ${y} ${x + w} ${y + tr - tr * k} ${x + w} ${y + tr}` : '',
		`L ${x + w} ${y + h - br}`,
		br
			? `C ${x + w} ${y + h - br + br * k} ${x + w - br + br * k} ${y + h} ${x + w - br} ${y + h}`
			: '',
		`L ${x + bl} ${y + h}`,
		bl ? `C ${x + bl - bl * k} ${y + h} ${x} ${y + h - bl + bl * k} ${x} ${y + h - bl}` : '',
		`L ${x} ${y + tl}`,
		tl ? `C ${x} ${y + tl - tl * k} ${x + tl - tl * k} ${y} ${x + tl} ${y}` : '',
		'Z',
	]
		.filter(Boolean)
		.join(' ');
}

/** One corner radius (`8px` or `8px 4px`); `undefined` when elliptical. */
export function cornerRadius(value: string): number | undefined {
	const parts = value.trim().split(/\s+/).map(Number.parseFloat);
	const [rx = 0, ry = rx] = parts;
	if (!Number.isFinite(rx) || !Number.isFinite(ry)) return undefined;
	if (value.includes('%')) return undefined;
	return Math.abs(rx - ry) < 0.5 ? rx : undefined;
}

const length = (value: string, size: number): number | undefined => {
	const parsed = Number.parseFloat(value);
	if (!Number.isFinite(parsed)) return undefined;
	return value.trim().endsWith('%') ? (parsed / 100) * size : parsed;
};

/**
 * A computed `clip-path` as path data in the element's border-box px, `null` for none, and
 * `undefined` for a value this cannot represent (`url()`, a shape with a reference box other
 * than the border box).
 */
export function clipPathData(
	value: string,
	width: number,
	height: number,
): { d: string; evenOdd: boolean } | null | undefined {
	let text = value.trim();
	if (!text || text === 'none') return null;
	text = text.replace(/\s+border-box$/, '');
	const call = /^([a-z]+)\((.*)\)$/s.exec(text);
	if (!call) return undefined;
	const [, kind, rawArgs] = call;
	let args = rawArgs!.trim();
	let evenOdd = false;
	const rule = /^(evenodd|nonzero)\s*,\s*/.exec(args);
	if (rule) {
		evenOdd = rule[1] === 'evenodd';
		args = args.slice(rule[0].length);
	}
	switch (kind) {
		case 'path': {
			const d = /^(['"])(.*)\1$/s.exec(args)?.[2];
			return d === undefined ? undefined : { d, evenOdd };
		}
		case 'polygon': {
			const points = args.split(',').map((pair) => pair.trim().split(/\s+/));
			const coords: string[] = [];
			for (const [index, pair] of points.entries()) {
				const x = length(pair[0] ?? '', width),
					y = length(pair[1] ?? '', height);
				if (x === undefined || y === undefined) return undefined;
				coords.push(`${index ? 'L' : 'M'} ${x} ${y}`);
			}
			return coords.length ? { d: `${coords.join(' ')} Z`, evenOdd } : undefined;
		}
		case 'inset': {
			if (/round/.test(args)) return undefined;
			const v = args.split(/\s+/).map((part, i) => length(part, i % 2 ? width : height));
			if (v.some((part) => part === undefined)) return undefined;
			const [top = 0, right = top, bottom = top, left = right] = v as number[];
			return {
				d: roundedRectPath(left, top, width - left - right, height - top - bottom),
				evenOdd,
			};
		}
		case 'circle':
		case 'ellipse': {
			const at = /^(.*?)\s*at\s+(\S+)\s+(\S+)$/.exec(args);
			const radii = (at ? at[1]! : args).trim().split(/\s+/).filter(Boolean);
			const cx = at ? length(at[2]!, width) : width / 2;
			const cy = at ? length(at[3]!, height) : height / 2;
			const diagonal = Math.hypot(width, height) / Math.SQRT2;
			const rx =
				kind === 'circle' ? length(radii[0] ?? '50%', diagonal) : length(radii[0] ?? '50%', width);
			const ry = kind === 'circle' ? rx : length(radii[1] ?? '50%', height);
			if ([cx, cy, rx, ry].some((n) => n === undefined || !Number.isFinite(n))) return undefined;
			const k = 0.5522847498;
			const [X, Y, RX, RY] = [cx!, cy!, rx!, ry!];
			return {
				d: [
					`M ${X + RX} ${Y}`,
					`C ${X + RX} ${Y + RY * k} ${X + RX * k} ${Y + RY} ${X} ${Y + RY}`,
					`C ${X - RX * k} ${Y + RY} ${X - RX} ${Y + RY * k} ${X - RX} ${Y}`,
					`C ${X - RX} ${Y - RY * k} ${X - RX * k} ${Y - RY} ${X} ${Y - RY}`,
					`C ${X + RX * k} ${Y - RY} ${X + RX} ${Y - RY * k} ${X + RX} ${Y}`,
					'Z',
				].join(' '),
				evenOdd,
			};
		}
		default:
			return undefined;
	}
}

/** Split a CSS list on top-level commas (not those inside parentheses). */
export function splitTopLevel(value: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let start = 0;
	for (let i = 0; i < value.length; i++) {
		const ch = value[i];
		if (ch === '(') depth++;
		else if (ch === ')') depth--;
		else if (ch === ',' && depth === 0) {
			out.push(value.slice(start, i).trim());
			start = i + 1;
		}
	}
	out.push(value.slice(start).trim());
	return out.filter(Boolean);
}

/**
 * One computed `linear-gradient(...)` as an axial shading over a `width` x `height` box, in
 * box px. `undefined` for anything else (radial, conic, repeating, hints, lengths in stops).
 */
export function linearGradientShading(
	value: string,
	width: number,
	height: number,
): { shading: PdfShading; opacity: number; variableOpacity: boolean } | undefined {
	const call = /^linear-gradient\((.*)\)$/s.exec(value.trim());
	if (!call || !(width > 0 && height > 0)) return undefined;
	const parts = splitTopLevel(call[1]!);
	let angle = 180;
	const head = parts[0] ?? '';
	if (/^-?[\d.]+(deg|turn|rad|grad)$/.test(head)) {
		const n = Number.parseFloat(head);
		angle = head.endsWith('turn')
			? n * 360
			: head.endsWith('grad')
				? n * 0.9
				: head.endsWith('rad')
					? (n * 180) / Math.PI
					: n;
		parts.shift();
	} else if (head.startsWith('to ')) {
		const sides = head.slice(3).split(/\s+/);
		const dx = sides.includes('right') ? 1 : sides.includes('left') ? -1 : 0;
		const dy = sides.includes('bottom') ? 1 : sides.includes('top') ? -1 : 0;
		// Corner directions point at the corner of this box.
		angle = (Math.atan2(dx * height, -dy * width) * 180) / Math.PI;
		if (!dx && !dy) return undefined;
		parts.shift();
	}
	const stops: { offset: number; r: number; g: number; b: number; a: number }[] = [];
	for (const part of parts) {
		const match = /^(.*?)(?:\s+(-?[\d.]+)%)?(?:\s+(-?[\d.]+)%)?$/s.exec(part);
		if (!match) return undefined;
		const color = cssColor(match[1]);
		if (color === undefined) return undefined;
		const c = color ?? { r: 0, g: 0, b: 0, a: 0 };
		const first = match[2] === undefined ? undefined : Number(match[2]) / 100;
		stops.push({ offset: first ?? Number.NaN, ...c });
		if (match[3] !== undefined) stops.push({ offset: Number(match[3]) / 100, ...c });
	}
	if (stops.length < 2) return undefined;
	// Distribute stops without a position, as CSS does.
	if (Number.isNaN(stops[0]!.offset)) stops[0]!.offset = 0;
	if (Number.isNaN(stops[stops.length - 1]!.offset)) stops[stops.length - 1]!.offset = 1;
	for (let i = 1; i < stops.length; i++) {
		if (!Number.isNaN(stops[i]!.offset)) continue;
		let j = i;
		while (Number.isNaN(stops[j]!.offset)) j++;
		const from = stops[i - 1]!.offset,
			to = stops[j]!.offset;
		for (let k = i; k < j; k++) stops[k]!.offset = from + ((to - from) * (k - i + 1)) / (j - i + 1);
	}
	for (let i = 1; i < stops.length; i++)
		stops[i]!.offset = Math.max(stops[i]!.offset, stops[i - 1]!.offset);
	const rad = (angle * Math.PI) / 180;
	const dx = Math.sin(rad),
		dy = -Math.cos(rad);
	const half = (Math.abs(width * dx) + Math.abs(height * dy)) / 2;
	const cx = width / 2,
		cy = height / 2;
	const opacities = stops.map((stop) => stop.a);
	return {
		shading: {
			kind: 'axial',
			coords: [cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half],
			stops: stops.map(({ offset, r, g, b }) => ({
				offset: Math.min(1, Math.max(0, offset)),
				r,
				g,
				b,
			})),
		},
		opacity: opacities.reduce((sum, a) => sum + a, 0) / opacities.length,
		variableOpacity: opacities.some((a) => Math.abs(a - opacities[0]!) > 0.01),
	};
}
