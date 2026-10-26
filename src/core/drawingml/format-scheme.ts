// Theme style matrices retain their source XML. Style references are one based; background
// fill references start at 1001. Missing or invalid indices never invent an Office default.
import { NS, buildXml, children, elements, first, parseXml, type XmlElement } from '../xml/index';
import { parseDrawingFill, parseDrawingLine } from './drawing-fill';
import { drawingFillXml } from './write-fill';
import { drawingLineXml } from './write-line';
import { drawingEffectsXml, parseDrawingEffects } from './drawing-effects';
import type { DrawingEffectList } from './effect-model';
import type { DrawingColor, DrawingFill, DrawingLine } from './types';
import { patchDrawingColorIn } from './patch-color';

export interface ThemeFillStyle {
	fill: DrawingFill;
	sourceXml?: string;
}
export interface ThemeEffectStyle {
	effects?: DrawingEffectList;
	sourceXml?: string;
}
export interface ThemeFormatScheme {
	name?: string;
	fillStyles: ThemeFillStyle[];
	lineStyles: DrawingLine[];
	effectStyles: ThemeEffectStyle[];
	backgroundFillStyles: ThemeFillStyle[];
	sourceXml?: string;
}

function parseFill(node: XmlElement): ThemeFillStyle {
	const doc = parseXml(`<a:spPr xmlns:a="${NS.a}"/>`);
	doc.documentElement.appendChild(doc.importNode(node, true));
	return {
		fill: parseDrawingFill(doc.documentElement) ?? { kind: 'unsupported', element: node.localName },
		sourceXml: buildXml(node),
	};
}

export function parseThemeFormatScheme(
	root: XmlElement | undefined,
): ThemeFormatScheme | undefined {
	if (!root || root.namespaceURI !== NS.a || root.localName !== 'fmtScheme') return undefined;
	const fills = (name: string) => {
		const list = first(root, name, NS.a);
		return list
			? elements(list)
					.filter((node) => node.namespaceURI === NS.a)
					.map(parseFill)
			: [];
	};
	const effects = first(root, 'effectStyleLst', NS.a);
	return {
		...(root.hasAttribute('name') ? { name: root.getAttribute('name')! } : {}),
		fillStyles: fills('fillStyleLst'),
		backgroundFillStyles: fills('bgFillStyleLst'),
		lineStyles: children(first(root, 'lnStyleLst', NS.a) ?? root, 'ln', NS.a).flatMap((node) => {
			const line = parseDrawingLine(node);
			return line ? [line] : [];
		}),
		effectStyles: effects
			? children(effects, 'effectStyle', NS.a).map((node) => {
					const list = parseDrawingEffects(first(node, 'effectLst', NS.a));
					return { ...(list ? { effects: list } : {}), sourceXml: buildXml(node) };
				})
			: [],
		sourceXml: buildXml(root),
	};
}

function fillXml(style: ThemeFillStyle): string {
	if (style.sourceXml) {
		const before = parseFill(parseXml(style.sourceXml).documentElement);
		if (JSON.stringify(before.fill) === JSON.stringify(style.fill)) return style.sourceXml;
		if (before.fill.kind === 'solid' && style.fill.kind === 'solid') {
			const node = parseXml(style.sourceXml).documentElement;
			patchDrawingColorIn(node, style.fill.color);
			return buildXml(node);
		}
	}
	const xml = drawingFillXml(style.fill);
	if (!xml) {
		if (style.sourceXml) return style.sourceXml;
		throw new Error('Cannot author this DrawingML theme fill');
	}
	return xml;
}

function effectXml(style: ThemeEffectStyle): string {
	const doc = parseXml(style.sourceXml ?? `<a:effectStyle xmlns:a="${NS.a}"/>`);
	const root = doc.documentElement;
	const old = first(root, 'effectLst', NS.a);
	if (style.effects) {
		const next = doc.importNode(parseXml(drawingEffectsXml(style.effects)).documentElement, true);
		if (old) root.replaceChild(next, old);
		else root.insertBefore(next, root.firstChild);
	} else if (old) root.removeChild(old);
	return buildXml(doc);
}

export function drawingThemeFormatSchemeXml(scheme: ThemeFormatScheme): string {
	const doc = parseXml(scheme.sourceXml ?? `<a:fmtScheme xmlns:a="${NS.a}"/>`);
	const root = doc.documentElement;
	if (root.namespaceURI !== NS.a || root.localName !== 'fmtScheme')
		throw new Error('Invalid DrawingML format scheme source');
	if (scheme.name === undefined) root.removeAttribute('name');
	else root.setAttribute('name', scheme.name);
	const lists = [
		['fillStyleLst', scheme.fillStyles.map(fillXml), undefined],
		['lnStyleLst', scheme.lineStyles.map(drawingLineXml), 'ln'],
		['effectStyleLst', scheme.effectStyles.map(effectXml), 'effectStyle'],
		['bgFillStyleLst', scheme.backgroundFillStyles.map(fillXml), undefined],
	] as const;
	for (let index = 0; index < lists.length; index++) {
		const [name, values, local] = lists[index]!;
		let list = first(root, name, NS.a);
		// Preserve absent lists in imported overrides; author all four lists for new matrices.
		if (!list && !values.length && scheme.sourceXml) continue;
		if (!list) {
			list = doc.createElementNS(NS.a, `a:${name}`);
			const following = lists.slice(index + 1).map(([next]) => next);
			root.insertBefore(
				list,
				elements(root).find(
					(node) =>
						following.some((next) => node.localName === next) || node.localName === 'extLst',
				) ?? null,
			);
		}
		const old = elements(list).filter(
			(node) => node.namespaceURI === NS.a && (!local || node.localName === local),
		);
		values.forEach((xml, offset) => {
			const wrapper = parseXml(`<w xmlns:a="${NS.a}">${xml}</w>`);
			const node = doc.importNode(elements(wrapper.documentElement)[0]!, true);
			if (old[offset]) list!.replaceChild(node, old[offset]!);
			else list!.appendChild(node);
		});
		for (const node of old.slice(values.length)) list.removeChild(node);
	}
	return buildXml(doc);
}

/** Substitutes phClr without flattening either the reference's theme token or its transforms. */
export function substituteDrawingPlaceholder(
	color: DrawingColor,
	placeholder?: DrawingColor,
): DrawingColor {
	if (color.kind !== 'scheme' || color.value !== 'phClr' || !placeholder)
		return structuredClone(color);
	return {
		...structuredClone(placeholder),
		transforms: [...structuredClone(placeholder.transforms), ...structuredClone(color.transforms)],
	};
}

function fillPlaceholder(fill: DrawingFill, placeholder?: DrawingColor): DrawingFill {
	if (fill.kind === 'solid')
		return { ...fill, color: substituteDrawingPlaceholder(fill.color, placeholder) };
	if (fill.kind === 'gradient')
		return {
			...structuredClone(fill),
			stops: fill.stops.map((stop) => ({
				...stop,
				color: substituteDrawingPlaceholder(stop.color, placeholder),
			})),
		};
	if (fill.kind === 'pattern')
		return {
			...fill,
			...(fill.foreground
				? { foreground: substituteDrawingPlaceholder(fill.foreground, placeholder) }
				: {}),
			...(fill.background
				? { background: substituteDrawingPlaceholder(fill.background, placeholder) }
				: {}),
		};
	return structuredClone(fill);
}

export function resolveThemeFillStyle(
	scheme: ThemeFormatScheme,
	index: number,
	placeholder?: DrawingColor,
): DrawingFill | undefined {
	if (!Number.isSafeInteger(index) || index < 1 || index === 1000) return undefined;
	const style =
		index > 1000 ? scheme.backgroundFillStyles[index - 1001] : scheme.fillStyles[index - 1];
	return style ? fillPlaceholder(style.fill, placeholder) : undefined;
}

export function resolveThemeLineStyle(
	scheme: ThemeFormatScheme,
	index: number,
	placeholder?: DrawingColor,
): DrawingLine | undefined {
	if (!Number.isSafeInteger(index) || index < 1) return undefined;
	const style = scheme.lineStyles[index - 1];
	return style
		? {
				...structuredClone(style),
				...(style.fill ? { fill: fillPlaceholder(style.fill, placeholder) } : {}),
			}
		: undefined;
}

export function resolveThemeEffectStyle(
	scheme: ThemeFormatScheme,
	index: number,
	placeholder?: DrawingColor,
): DrawingEffectList | undefined {
	if (!Number.isSafeInteger(index) || index < 1) return undefined;
	const list = scheme.effectStyles[index - 1]?.effects;
	if (!list) return undefined;
	return {
		...structuredClone(list),
		effects: list.effects.map((effect) => ({
			...structuredClone(effect),
			...('color' in effect && effect.color
				? { color: substituteDrawingPlaceholder(effect.color, placeholder) }
				: {}),
			...(effect.kind === 'fillOverlay' && effect.fill
				? { fill: fillPlaceholder(effect.fill, placeholder) }
				: {}),
		})),
	};
}
