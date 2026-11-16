// CT_EffectList codec. Source-backed edits change only the requested properties, preserving
// extension XML, invalid unmodelled attributes and authored scheme colours in every other field.
import { NS, buildXml, elements, parseXml, type XmlElement } from '../xml/index';
import { fillElementOf, parseDrawingFill } from './drawing-fill';
import { drawingFillXml } from './write-fill';
import type { DrawingEffect, DrawingEffectList } from './effect-model';
import {
	DRAWING_EFFECT_BOOLEANS,
	DRAWING_EFFECT_COLORED,
	DRAWING_EFFECT_KINDS,
	DRAWING_EFFECT_NUMBERS,
	DRAWING_EFFECT_STRINGS,
	readDrawingEffect,
	type DrawingEffectKind,
} from './effect-read';
import { DOM_DRAWING_VIEW } from './node-view';
import { patchDrawingColorIn } from './patch-color';

const KINDS = DRAWING_EFFECT_KINDS;
type KnownKind = DrawingEffectKind;
const TAGS = Object.fromEntries(Object.entries(KINDS).map(([tag, kind]) => [kind, tag])) as Record<
	KnownKind,
	string
>;
const NUMBERS = DRAWING_EFFECT_NUMBERS;
const STRINGS = DRAWING_EFFECT_STRINGS;
const BOOLEANS = DRAWING_EFFECT_BOOLEANS;
const COLORED = DRAWING_EFFECT_COLORED;
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function fragment(xml: string): XmlElement {
	return elements(parseXml(`<w xmlns:a="${NS.a}">${xml}</w>`).documentElement)[0]!;
}

/** Reads one effect. No defaults are synthesized and unsupported nodes remain editable opaque data. */
export function parseDrawingEffect(node: XmlElement): DrawingEffect {
	const sourceXml = buildXml(node);
	const fields =
		node.namespaceURI === NS.a
			? readDrawingEffect(DOM_DRAWING_VIEW, node.localName, node, parseDrawingFill)
			: undefined;
	if (!fields) return { kind: 'unsupported', sourceXml };
	const { kind, ...rest } = fields;
	return { kind, sourceXml, ...rest } as DrawingEffect;
}

/** Reads an a:effectLst. Absence is inheritance; an empty list is explicit no effects. */
export function parseDrawingEffects(list: XmlElement | undefined): DrawingEffectList | undefined {
	if (!list || list.namespaceURI !== NS.a || list.localName !== 'effectLst') return undefined;
	return { effects: elements(list).map(parseDrawingEffect), sourceXml: buildXml(list) };
}

/** Writes one effect, using its original node as a preservation template. */
export function drawingEffectXml(effect: DrawingEffect): string {
	if (effect.kind === 'unsupported') return effect.sourceXml;
	const doc = parseXml(effect.sourceXml ?? `<a:${TAGS[effect.kind]} xmlns:a="${NS.a}"/>`);
	const root = doc.documentElement;
	if (root.namespaceURI !== NS.a || root.localName !== TAGS[effect.kind])
		throw new Error('Invalid DrawingML effect source XML');
	const before = effect.sourceXml ? parseDrawingEffect(root) : undefined;
	const current = effect as unknown as Record<string, unknown>;
	const baseline = before as unknown as Record<string, unknown> | undefined;
	const changed = (key: string) => !baseline || !same(baseline[key], current[key]);
	for (const [key, attribute] of Object.entries({
		...NUMBERS[effect.kind],
		...STRINGS[effect.kind],
		...BOOLEANS[effect.kind],
	})) {
		if (!changed(key)) continue;
		const value = current[key];
		if (value === undefined) root.removeAttribute(attribute);
		else {
			if (typeof value === 'number' && !Number.isFinite(value))
				throw new Error('Invalid DrawingML effect value');
			root.setAttribute(
				attribute,
				typeof value === 'boolean' ? (value ? '1' : '0') : String(value),
			);
		}
	}
	if (COLORED.has(effect.kind) && changed('color')) {
		const color = 'color' in effect ? effect.color : undefined;
		if (color) patchDrawingColorIn(root, color);
		else {
			for (const child of elements(root))
				if (
					child.namespaceURI === NS.a &&
					['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'].includes(
						child.localName,
					)
				)
					root.removeChild(child);
		}
	}
	if (effect.kind === 'fillOverlay' && changed('fill')) {
		const xml = effect.fill && drawingFillXml(effect.fill);
		if (effect.fill && !xml) throw new Error('Cannot author this DrawingML effect fill');
		if (xml || !effect.fill) {
			const fill = fillElementOf(root);
			if (fill) root.removeChild(fill);
			if (xml) root.insertBefore(doc.importNode(fragment(xml), true), root.firstChild);
		}
	}
	return buildXml(doc);
}

/** Replaces the effect-list projection while retaining the list's attributes and non-element nodes. */
export function patchDrawingEffects(list: XmlElement, value: DrawingEffectList): void {
	if (list.namespaceURI !== NS.a || list.localName !== 'effectLst')
		throw new Error('Invalid DrawingML effect list');
	const original = elements(list);
	value.effects.forEach((effect, index) => {
		const next = list.ownerDocument.importNode(
			parseXml(drawingEffectXml(effect)).documentElement,
			true,
		);
		if (original[index]) list.replaceChild(next, original[index]!);
		else list.appendChild(next);
	});
	for (const old of original.slice(value.effects.length)) list.removeChild(old);
}

export function drawingEffectsXml(value: DrawingEffectList): string {
	const doc = parseXml(value.sourceXml ?? `<a:effectLst xmlns:a="${NS.a}"/>`);
	patchDrawingEffects(doc.documentElement, value);
	return buildXml(doc);
}
