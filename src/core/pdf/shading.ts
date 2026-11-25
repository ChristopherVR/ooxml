/**
 * Reads PDF shadings (Type 2 axial, Type 3 radial) and patterns (Type 1 tiling, Type 2 shading)
 * into format-neutral gradient and pattern fills for vector drawings.
 */
import type { PdfDocument } from './document';
import { EOF, PdfKeyword, PdfLexer } from './lexer';
import { apply, multiply, type Matrix } from './matrix';
import { isDict, isName, isNumber, PdfStream, type PdfDict, type PdfValue } from './objects';
import type { PdfGradientFill, PdfPatternFill } from './page-content';

const clampByte = (value: number) => Math.max(0, Math.min(255, Math.round(value * 255)));
const hexByte = (value: number) => clampByte(value).toString(16).padStart(2, '0');

function componentsToHex(values: number[]): string {
	if (values.length === 1) {
		const g = hexByte(values[0]!);
		return `#${g}${g}${g}`;
	}
	if (values.length === 4) {
		const [c, m, y, k] = values as [number, number, number, number];
		return `#${hexByte((1 - c) * (1 - k))}${hexByte((1 - m) * (1 - k))}${hexByte((1 - y) * (1 - k))}`;
	}
	const [r = 0, g = 0, b = 0] = values;
	return `#${hexByte(r)}${hexByte(g)}${hexByte(b)}`;
}

async function numberArray(document: PdfDocument, value: PdfValue | undefined): Promise<number[]> {
	const resolved = await document.resolve(value);
	if (!Array.isArray(resolved)) return [];
	const out: number[] = [];
	for (const item of resolved) {
		const n = await document.resolve(item);
		if (isNumber(n)) out.push(n);
	}
	return out;
}

async function evalFunction(
	document: PdfDocument,
	fnValue: PdfValue | undefined,
	t: number,
): Promise<number[] | undefined> {
	const resolved = await document.resolve(fnValue);
	const dict =
		resolved instanceof PdfStream ? resolved.dict : isDict(resolved) ? resolved : undefined;
	if (!dict) return undefined;
	const type = await document.resolve(dict.get('FunctionType'));
	if (type === 2) {
		const c0 = await numberArray(document, dict.get('C0'));
		const c1 = await numberArray(document, dict.get('C1'));
		const nVal = await document.resolve(dict.get('N'));
		const n = isNumber(nVal) ? nVal : 1;
		const start = c0.length ? c0 : [0];
		const end = c1.length ? c1 : [1];
		const k = Math.pow(Math.max(0, Math.min(1, t)), n);
		return start.map((v, i) => v + k * ((end[i] ?? v) - v));
	}
	if (type === 3) {
		const funcs = await document.resolve(dict.get('Functions'));
		if (!Array.isArray(funcs) || !funcs.length) return undefined;
		const bounds = await numberArray(document, dict.get('Bounds'));
		const encode = await numberArray(document, dict.get('Encode'));
		const domain = await numberArray(document, dict.get('Domain'));
		const d0 = domain[0] ?? 0;
		const d1 = domain[1] ?? 1;
		const x = d0 + t * (d1 - d0);
		let index = 0;
		while (index < bounds.length && x > bounds[index]!) index++;
		const subStart = index === 0 ? d0 : bounds[index - 1]!;
		const subEnd = index < bounds.length ? bounds[index]! : d1;
		const span = subEnd - subStart;
		const local = span > 0 ? (x - subStart) / span : 0;
		const e0 = encode[index * 2] ?? 0;
		const e1 = encode[index * 2 + 1] ?? 1;
		return evalFunction(document, funcs[index], e0 + local * (e1 - e0));
	}
	return undefined;
}

async function readStops(
	document: PdfDocument,
	fnValue: PdfValue | undefined,
): Promise<{ pos: number; color: string }[]> {
	const resolved = await document.resolve(fnValue);
	const dict =
		resolved instanceof PdfStream ? resolved.dict : isDict(resolved) ? resolved : undefined;
	const positions = new Set<number>([0, 1]);
	if (dict) {
		const type = await document.resolve(dict.get('FunctionType'));
		if (type === 3) {
			const bounds = await numberArray(document, dict.get('Bounds'));
			const domain = await numberArray(document, dict.get('Domain'));
			const d0 = domain[0] ?? 0;
			const d1 = domain[1] ?? 1;
			const span = d1 - d0 || 1;
			for (const b of bounds) positions.add(Math.max(0, Math.min(1, (b - d0) / span)));
		}
	}
	const sorted = [...positions].sort((a, b) => a - b);
	const stops: { pos: number; color: string }[] = [];
	for (const pos of sorted) {
		const values = await evalFunction(document, fnValue, pos);
		if (values?.length) stops.push({ pos, color: componentsToHex(values) });
	}
	return stops;
}

/** Reads a PDF Shading dictionary (Type 2 axial or Type 3 radial) into a gradient fill. */
export async function readPdfShading(
	document: PdfDocument,
	shadingValue: PdfValue | undefined,
	ctm: Matrix,
	page: Matrix,
): Promise<PdfGradientFill | undefined> {
	const dict = await document.dict(shadingValue);
	if (!dict) return undefined;
	const type = await document.resolve(dict.get('ShadingType'));
	if (type !== 2 && type !== 3) return undefined;
	const stops = await readStops(document, dict.get('Function'));
	if (stops.length < 2) return undefined;
	if (type === 3) return { kind: 'radial', stops };
	const coords = await numberArray(document, dict.get('Coords'));
	let angle = 0;
	if (coords.length >= 4) {
		const m = multiply(ctm, page);
		const [x0, y0] = apply(m, coords[0]!, coords[1]!);
		const [x1, y1] = apply(m, coords[2]!, coords[3]!);
		const deg = (Math.atan2(y1 - y0, x1 - x0) * 180) / Math.PI;
		angle = ((deg % 360) + 360) % 360;
	}
	return { kind: 'linear', angle, stops };
}

/** Reads a PDF Pattern resource (Type 2 shading pattern or Type 1 tiling pattern). */
export async function readPdfPattern(
	document: PdfDocument,
	resources: PdfDict,
	name: string,
	ctm: Matrix,
	page: Matrix,
	fallbackColor: string,
): Promise<{ gradient?: PdfGradientFill; patternFill?: PdfPatternFill } | undefined> {
	const patterns = await document.dict(resources.get('Pattern'));
	const raw = await document.resolve(patterns?.get(name));
	const dict = raw instanceof PdfStream ? raw.dict : isDict(raw) ? raw : undefined;
	if (!dict) return undefined;
	const type = await document.resolve(dict.get('PatternType'));
	if (type === 2) {
		const matVals = await numberArray(document, dict.get('Matrix'));
		const patMat: Matrix = matVals.length === 6 ? (matVals as Matrix) : [1, 0, 0, 1, 0, 0];
		const gradient = await readPdfShading(
			document,
			dict.get('Shading'),
			multiply(patMat, ctm),
			page,
		);
		return gradient ? { gradient } : undefined;
	}
	if (type === 1) {
		let fg = fallbackColor;
		const paintType = await document.resolve(dict.get('PaintType'));
		if (paintType === 1 && raw instanceof PdfStream) {
			try {
				const bytes = await document.streamBytes(raw);
				const lexer = new PdfLexer(bytes);
				const ops: PdfValue[] = [];
				for (let i = 0; i < 500; i++) {
					const tok = lexer.object();
					if (tok === EOF) break;
					if (!(tok instanceof PdfKeyword)) {
						ops.push(tok);
						continue;
					}
					if (['g', 'rg', 'k', 'G', 'RG', 'K'].includes(tok.word)) {
						const nums = ops.filter(isNumber);
						if (nums.length) {
							fg = componentsToHex(nums);
							break;
						}
					}
					ops.length = 0;
				}
			} catch {
				// Keep fallback colour if the tiling stream cannot be inspected.
			}
		}
		const nameLower = name.toLowerCase();
		const preset = nameLower.includes('diag')
			? 'ltDnDiag'
			: nameLower.includes('horz')
				? 'ltHorz'
				: nameLower.includes('vert')
					? 'ltVert'
					: nameLower.includes('cross')
						? 'smGrid'
						: 'pct25';
		const isColor = isName(await document.resolve(dict.get('Type'))) || paintType !== undefined;
		return isColor ? { patternFill: { preset, fg, bg: '#ffffff' } } : undefined;
	}
	return undefined;
}
