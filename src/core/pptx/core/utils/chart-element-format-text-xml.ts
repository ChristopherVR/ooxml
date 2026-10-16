/**
 * Writes and reads the Text Options of a chart element (text fill, outline,
 * shadow and glow, custom angle, direction and vertical alignment) on its
 * `c:txPr`, or on a title's own rich text (`c:tx/c:rich`) when it has one:
 * the `a:bodyPr` attributes, every paragraph's `a:defRPr` and every run's
 * `a:rPr`.
 *
 * @module utils/chart-element-format-text-xml
 */
import type {
	PptxChartElementEffects,
	PptxChartElementText,
	PptxChartTextFill,
	PptxChartTextOutline,
	XmlObject,
} from '../types';
import { applyChartEffects } from './chart-element-format-effects-xml';
import { FILL_CHOICE, LN_FILL_CHOICE, applyChartPaint } from './chart-element-format-fill-xml';
import type { LocalName } from './chart-element-format-xml-util';
import {
	asObject,
	keyOf,
	pt,
	removeChildren,
	replaceEntries,
	setOrdered,
} from './chart-element-format-xml-util';

const RPR_ORDER = [
	'ln',
	...FILL_CHOICE,
	'effectLst',
	'effectDag',
	'highlight',
	'uLnTx',
	'uLn',
	'uFillTx',
	'uFill',
	'latin',
	'ea',
	'cs',
	'sym',
	'hlinkClick',
	'hlinkMouseOver',
	'rtl',
	'extLst',
];
const PPR_ORDER = [
	'lnSpc',
	'spcBef',
	'spcAft',
	'buClrTx',
	'buClr',
	'buSzTx',
	'buSzPct',
	'buSzPts',
	'buFontTx',
	'buFont',
	'buNone',
	'buAutoNum',
	'buChar',
	'buBlip',
	'tabLst',
	'defRPr',
	'extLst',
];
const LN_ORDER = [...LN_FILL_CHOICE, 'prstDash', 'custDash', 'round', 'bevel', 'miter'];

function list(value: unknown): XmlObject[] {
	const items = Array.isArray(value) ? value : value === undefined ? [] : [value];
	return items.map((item) => asObject(item) ?? {});
}

function applyOutline(rPr: XmlObject, outline: PptxChartTextOutline, localName: LocalName): void {
	if (outline.kind === 'automatic') {
		removeChildren(rPr, ['ln'], localName);
		return;
	}
	const key = keyOf(rPr, 'ln', localName);
	const ln: XmlObject = { ...(asObject(key ? rPr[key] : undefined) ?? {}) };
	applyChartPaint(ln, outline, LN_FILL_CHOICE, LN_ORDER, localName);
	if (outline.kind === 'solid' && outline.width !== undefined) ln['@_w'] = pt(outline.width);
	if (outline.kind === 'solid' && outline.dash !== undefined) {
		removeChildren(ln, ['custDash'], localName);
		setOrdered(ln, 'prstDash', { '@_val': outline.dash }, LN_ORDER, localName);
	}
	setOrdered(rPr, 'ln', ln, RPR_ORDER, localName);
}

/** Apply the run-level text settings to one `a:defRPr` / `a:rPr`. */
function applyRunProps(rPr: XmlObject, text: PptxChartElementText, localName: LocalName): void {
	if (text.outline) applyOutline(rPr, text.outline, localName);
	if (text.fill) applyChartPaint(rPr, text.fill, FILL_CHOICE, RPR_ORDER, localName);
	if (text.effects) applyChartEffects(rPr, text.effects, RPR_ORDER, localName);
}

function applyBodyProps(bodyPr: XmlObject, text: PptxChartElementText): void {
	if (text.rotation !== undefined) {
		if (text.rotation === null) delete bodyPr['@_rot'];
		else bodyPr['@_rot'] = String(Math.round(Math.max(-90, Math.min(90, text.rotation)) * 60000));
	}
	if (text.direction !== undefined) {
		if (text.direction === null) delete bodyPr['@_vert'];
		else bodyPr['@_vert'] = text.direction;
	}
	if (text.anchor !== undefined) {
		if (text.anchor === null) delete bodyPr['@_anchor'];
		else bodyPr['@_anchor'] = text.anchor;
	}
}

const hasRunProps = (text: PptxChartElementText) =>
	Boolean(text.fill || text.outline || text.effects);

/** Apply `text` to a text body (`c:txPr` or `c:rich`) in place. */
function applyToBody(body: XmlObject, text: PptxChartElementText, localName: LocalName): void {
	const bodyKey = keyOf(body, 'bodyPr', localName) ?? 'a:bodyPr';
	const bodyPr: XmlObject = { ...(asObject(body[bodyKey]) ?? {}) };
	applyBodyProps(bodyPr, text);
	body[bodyKey] = bodyPr;
	if (!hasRunProps(text)) {
		return;
	}
	const pKey = keyOf(body, 'p', localName) ?? 'a:p';
	const paragraphs = list(body[pKey]);
	if (paragraphs.length === 0) paragraphs.push({});
	const next = paragraphs.map((paragraph) => {
		const p: XmlObject = { ...paragraph };
		const existingPPr = keyOf(p, 'pPr', localName);
		const pPrKey = existingPPr ?? 'a:pPr';
		const pPr: XmlObject = { ...(asObject(p[pPrKey]) ?? {}) };
		const defKey = keyOf(pPr, 'defRPr', localName);
		const rKey = keyOf(p, 'r', localName);
		// A rich title paragraph without its own defaults keeps none: an added `a:defRPr`
		// (even one without a size) replaces the chart style's title size and weight.
		if (defKey || !rKey) {
			const defRPr: XmlObject = { ...(asObject(defKey ? pPr[defKey] : undefined) ?? {}) };
			applyRunProps(defRPr, text, localName);
			setOrdered(pPr, 'defRPr', defRPr, PPR_ORDER, localName);
		}
		if (rKey) {
			const runs = list(p[rKey]).map((run) => {
				const r: XmlObject = { ...run };
				const rPrKey = keyOf(r, 'rPr', localName);
				const rPr: XmlObject = { ...(asObject(rPrKey ? r[rPrKey] : undefined) ?? {}) };
				applyRunProps(rPr, text, localName);
				// `a:rPr` comes before `a:t`.
				replaceEntries(r, [
					[rPrKey ?? 'a:rPr', rPr],
					...Object.entries(r).filter(([k]) => k !== rPrKey),
				]);
				return r;
			});
			p[rKey] = runs.length === 1 ? runs[0]! : runs;
		}
		// `a:pPr` is the paragraph's first child.
		if (existingPPr || Object.keys(pPr).length > 0) {
			replaceEntries(p, [[pPrKey, pPr], ...Object.entries(p).filter(([k]) => k !== pPrKey)]);
		}
		return p;
	});
	body[pKey] = next.length === 1 ? next[0]! : next;
}

/** A fresh `c:txPr` body. */
function newTextBody(): XmlObject {
	return {
		'a:bodyPr': {},
		'a:lstStyle': {},
		'a:p': { 'a:pPr': { 'a:defRPr': {} }, 'a:endParaRPr': { '@_lang': 'en-US' } },
	};
}

/**
 * Apply `text` to `owner`: its rich title text when it has one, else its
 * `c:txPr`, inserted before the first child named in `after` when missing.
 */
export function applyChartTextFormat(
	owner: XmlObject,
	text: PptxChartElementText,
	after: readonly string[],
	localName: LocalName,
): void {
	const tx = asObject(owner[keyOf(owner, 'tx', localName) ?? '']);
	const rich = tx ? asObject(tx[keyOf(tx, 'rich', localName) ?? '']) : undefined;
	if (rich) {
		applyToBody(rich, text, localName);
		return;
	}
	const key = keyOf(owner, 'txPr', localName);
	const body: XmlObject = key ? { ...(asObject(owner[key]) ?? {}) } : newTextBody();
	applyToBody(body, text, localName);
	if (key) {
		owner[key] = body;
		return;
	}
	const entries = Object.entries(owner);
	const at = entries.findIndex(([k]) => !k.startsWith('@_') && after.includes(localName(k)));
	entries.splice(at < 0 ? entries.length : at, 0, ['c:txPr', body]);
	replaceEntries(owner, entries);
}

// ── Reading ─────────────────────────────────────────────────────────────────

type ReadEffects = (list: XmlObject) => PptxChartElementEffects | undefined;
type ReadPaint = (node: XmlObject) => PptxChartTextFill | undefined;

/** The first run's properties, else the first paragraph's `a:defRPr`. */
function firstRunProps(body: XmlObject, localName: LocalName): XmlObject | undefined {
	const p = list(body[keyOf(body, 'p', localName) ?? ''])[0];
	if (!p) return undefined;
	const run = list(p[keyOf(p, 'r', localName) ?? ''])[0];
	const rPr = run ? asObject(run[keyOf(run, 'rPr', localName) ?? '']) : undefined;
	const pPr = asObject(p[keyOf(p, 'pPr', localName) ?? '']);
	const def = pPr ? asObject(pPr[keyOf(pPr, 'defRPr', localName) ?? '']) : undefined;
	if (!rPr) return def;
	if (!def) return rPr;
	// A run that sets no fill of its own paints with the paragraph default's.
	return { ...def, ...rPr };
}

/** The Text Options `owner` carries, or `undefined` when it has no text body. */
export function readChartTextFormat(
	owner: XmlObject | undefined,
	localName: LocalName,
	readPaint: ReadPaint,
	readEffects: ReadEffects,
): PptxChartElementText | undefined {
	if (!owner) return undefined;
	const tx = asObject(owner[keyOf(owner, 'tx', localName) ?? '']);
	const body =
		(tx ? asObject(tx[keyOf(tx, 'rich', localName) ?? '']) : undefined) ??
		asObject(owner[keyOf(owner, 'txPr', localName) ?? '']);
	if (!body) return undefined;
	const out: PptxChartElementText = {};
	const bodyPr = asObject(body[keyOf(body, 'bodyPr', localName) ?? '']) ?? {};
	const rot = Number(bodyPr['@_rot']);
	// PowerPoint writes `rot="-60000000"` for "automatic"; only a real angle (-90..90) is custom.
	if (bodyPr['@_rot'] !== undefined && Number.isFinite(rot) && Math.abs(rot) <= 5400000)
		out.rotation = Math.round(rot / 600) / 100;
	const vert = bodyPr['@_vert'];
	if (vert === 'horz' || vert === 'vert' || vert === 'vert270' || vert === 'wordArtVert')
		out.direction = vert;
	const anchor = bodyPr['@_anchor'];
	if (anchor === 't' || anchor === 'ctr' || anchor === 'b') out.anchor = anchor;
	const rPr = firstRunProps(body, localName);
	if (rPr) {
		const fill = readPaint(rPr);
		if (fill) out.fill = fill;
		const ln = asObject(rPr[keyOf(rPr, 'ln', localName) ?? '']);
		if (ln) {
			const paint = readPaint(ln);
			const outline: PptxChartTextOutline =
				paint?.kind === 'none'
					? { kind: 'none' }
					: paint?.kind === 'solid'
						? {
								kind: 'solid',
								...(paint.color ? { color: paint.color } : {}),
								...(paint.transparency !== undefined ? { transparency: paint.transparency } : {}),
							}
						: { kind: 'automatic' };
			const w = Number(ln['@_w']);
			if (outline.kind === 'solid' && Number.isFinite(w))
				outline.width = Math.round((w / 12700) * 100) / 100;
			const dash = asObject(ln[keyOf(ln, 'prstDash', localName) ?? ''])?.['@_val'];
			if (outline.kind === 'solid' && typeof dash === 'string') outline.dash = dash;
			out.outline = outline;
		}
		const effectLst = asObject(rPr[keyOf(rPr, 'effectLst', localName) ?? '']);
		const effects = effectLst ? readEffects(effectLst) : undefined;
		if (effects && (effects.shadow || effects.glow)) {
			out.effects = {
				...(effects.shadow ? { shadow: effects.shadow } : {}),
				...(effects.glow ? { glow: effects.glow } : {}),
			};
		}
	}
	return Object.keys(out).length > 0 ? out : undefined;
}
