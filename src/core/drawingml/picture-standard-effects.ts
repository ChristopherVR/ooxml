import { NS, elements, first, type XmlElement } from '../xml/index';
import { parseDrawingColor, parseDrawingColorIn } from './drawing-color';
import { patchDrawingColorIn } from './patch-color';
import type { DrawingColor } from './types';
export type PictureStandardEffectName =
	| 'biLevel'
	| 'alphaInv'
	| 'alphaCeiling'
	| 'alphaFloor'
	| 'alphaRepl'
	| 'alphaBiLevel'
	| 'alphaMod'
	| 'duotone'
	| 'clrRepl'
	| 'clrChange'
	| 'hsl'
	| 'tint'
	| 'blur';
export interface DrawingStandardPictureEffect {
	kind: 'standard';
	name: PictureStandardEffectName;
	values?: Partial<
		Record<'threshold' | 'amount' | 'hue' | 'saturation' | 'luminance' | 'radiusEmu', number>
	>;
	colors?: DrawingColor[];
	useAlpha?: boolean;
	grow?: boolean;
	sourceXml?: string;
}
const attrs: Partial<Record<PictureStandardEffectName, Record<string, string>>> = {
	biLevel: { threshold: 'thresh' },
	alphaRepl: { amount: 'a' },
	alphaBiLevel: { threshold: 'thresh' },
	hsl: { hue: 'hue', saturation: 'sat', luminance: 'lum' },
	tint: { hue: 'hue', amount: 'amt' },
	blur: { radiusEmu: 'rad' },
};
export function parseStandardPictureEffect(
	node: XmlElement,
): DrawingStandardPictureEffect | undefined {
	const name = node.localName as PictureStandardEffectName;
	if (
		![
			'biLevel',
			'alphaInv',
			'alphaCeiling',
			'alphaFloor',
			'alphaRepl',
			'alphaBiLevel',
			'alphaMod',
			'duotone',
			'clrRepl',
			'clrChange',
			'hsl',
			'tint',
			'blur',
		].includes(name)
	)
		return undefined;
	const values: NonNullable<DrawingStandardPictureEffect['values']> = {};
	for (const [key, attribute] of Object.entries(attrs[name] ?? {})) {
		const raw = node.getAttribute(attribute);
		if (raw === null) continue;
		const percent = raw.endsWith('%');
		const n =
			Number(percent ? raw.slice(0, -1) : raw) /
			(key === 'hue' ? 60000 : key === 'radiusEmu' ? 1 : percent ? 100 : 100000);
		if (
			!Number.isFinite(n) ||
			(key === 'hue' && (n < 0 || n >= 360)) ||
			(key === 'threshold' && (n < 0 || n > 1)) ||
			(key === 'radiusEmu' && (!Number.isSafeInteger(n) || n < 0)) ||
			(['saturation', 'luminance', 'amount'].includes(key) &&
				name !== 'alphaMod' &&
				(n < (name === 'alphaRepl' ? 0 : -1) || n > 1))
		)
			return undefined;
		values[key as keyof typeof values] = n;
	}
	if (['biLevel', 'alphaBiLevel'].includes(name) && values.threshold === undefined)
		return undefined;
	if (name === 'alphaRepl' && values.amount === undefined) return undefined;
	if (name === 'alphaMod') {
		const cont = first(node, 'cont', NS.a);
		const nested = cont ? elements(cont) : [];
		if (nested.length !== 1 || nested[0]?.localName !== 'alphaModFix') return undefined;
		const raw = nested[0].getAttribute('amt') ?? '100000';
		const n =
			Number(raw.endsWith('%') ? raw.slice(0, -1) : raw) / (raw.endsWith('%') ? 100 : 100000);
		if (!Number.isFinite(n) || n < 0) return undefined;
		values.amount = n;
	}
	let colors: DrawingColor[] = [];
	if (name === 'duotone') {
		const candidates = elements(node).filter(
			(item) => item.namespaceURI === NS.a && /Clr$/.test(item.localName),
		);
		colors = candidates
			.map((item) => parseDrawingColor(item))
			.filter((item): item is DrawingColor => Boolean(item));
		if (colors.length !== 2) return undefined;
	}
	if (name === 'clrRepl') {
		const color = parseDrawingColorIn(node);
		if (!color) return undefined;
		colors = [color];
	}
	if (name === 'clrChange') {
		const from = parseDrawingColorIn(first(node, 'clrFrom', NS.a)),
			to = parseDrawingColorIn(first(node, 'clrTo', NS.a));
		if (!from || !to) return undefined;
		colors = [from, to];
	}
	if (name === 'alphaInv') {
		const color = parseDrawingColorIn(node);
		if (color) colors = [color];
	}
	return {
		kind: 'standard',
		name,
		...(Object.keys(values).length ? { values } : {}),
		...(colors.length ? { colors } : {}),
		...(name === 'clrChange' && node.hasAttribute('useA')
			? { useAlpha: !['0', 'false'].includes(node.getAttribute('useA') ?? '') }
			: {}),
		...(name === 'blur' && node.hasAttribute('grow')
			? { grow: !['0', 'false'].includes(node.getAttribute('grow') ?? '') }
			: {}),
	};
}
export function applyStandardPictureEffect(
	node: XmlElement,
	effect: DrawingStandardPictureEffect,
): void {
	const doc = node.ownerDocument,
		before = parseStandardPictureEffect(node);
	for (const [key, value] of Object.entries(effect.values ?? {})) {
		if (
			!Number.isFinite(value) ||
			(key === 'hue' && (value < 0 || value >= 360)) ||
			(key === 'threshold' && (value < 0 || value > 1)) ||
			(key === 'radiusEmu' && (!Number.isSafeInteger(value) || value < 0)) ||
			(key === 'amount' && effect.name === 'alphaMod' && value < 0) ||
			(['saturation', 'luminance', 'amount'].includes(key) &&
				effect.name !== 'alphaMod' &&
				(value < (effect.name === 'alphaRepl' ? 0 : -1) || value > 1))
		)
			throw new RangeError('Invalid picture effect values');
	}
	const count = effect.colors?.length ?? 0;
	if (
		(['duotone', 'clrChange'].includes(effect.name) && count !== 2) ||
		(effect.name === 'clrRepl' && count !== 1) ||
		(effect.name === 'alphaInv' && count > 1)
	)
		throw new RangeError('Invalid picture effect colours');
	for (const [key, attribute] of Object.entries(attrs[effect.name] ?? {})) {
		const value = effect.values?.[key as keyof NonNullable<typeof effect.values>];
		if (value === before?.values?.[key as keyof NonNullable<typeof effect.values>]) continue;
		if (value === undefined) node.removeAttribute(attribute);
		else
			node.setAttribute(
				attribute,
				String(Math.round(value * (key === 'hue' ? 60000 : key === 'radiusEmu' ? 1 : 100000))),
			);
	}
	if (effect.name === 'alphaMod' && effect.values?.amount !== before?.values?.amount) {
		const cont = first(node, 'cont', NS.a) ?? doc.createElementNS(NS.a, 'a:cont');
		const fix = first(cont, 'alphaModFix', NS.a) ?? doc.createElementNS(NS.a, 'a:alphaModFix');
		fix.setAttribute('amt', String(Math.round((effect.values?.amount ?? 1) * 100000)));
		if (!fix.parentNode) cont.appendChild(fix);
		if (!cont.parentNode) node.appendChild(cont);
	}
	if (effect.useAlpha !== before?.useAlpha) {
		if (effect.useAlpha === undefined) node.removeAttribute('useA');
		else node.setAttribute('useA', effect.useAlpha ? '1' : '0');
	}
	if (effect.grow !== before?.grow) {
		if (effect.grow === undefined) node.removeAttribute('grow');
		else node.setAttribute('grow', effect.grow ? '1' : '0');
	}
	if (JSON.stringify(effect.colors) !== JSON.stringify(before?.colors)) {
		if (effect.name === 'duotone') {
			const old = elements(node).filter(
				(item) => item.namespaceURI === NS.a && /Clr$/.test(item.localName),
			);
			for (let i = 0; i < (effect.colors?.length ?? 0); i++) {
				const holder = doc.createElementNS(NS.a, 'a:holder');
				if (old[i]) holder.appendChild(doc.importNode(old[i]!, true));
				patchDrawingColorIn(holder, effect.colors![i]!);
				const next = elements(holder)[0]!;
				if (old[i]) node.replaceChild(next, old[i]!);
				else node.insertBefore(next, first(node, 'extLst', NS.a) ?? null);
			}
		} else if (effect.name === 'clrChange') {
			for (const [i, name] of ['clrFrom', 'clrTo'].entries()) {
				const holder = first(node, name, NS.a) ?? doc.createElementNS(NS.a, `a:${name}`);
				if (effect.colors?.[i]) patchDrawingColorIn(holder, effect.colors[i]!);
				if (!holder.parentNode) node.insertBefore(holder, first(node, 'extLst', NS.a) ?? null);
			}
		} else if (effect.colors?.[0]) patchDrawingColorIn(node, effect.colors[0]);
	}
	if (!parseStandardPictureEffect(node)) throw new RangeError('Invalid picture effect values');
}
