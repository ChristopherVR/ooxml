// DOM parsers for the three definition parts of a diagram: colours (`dgm:colorsDef`), quick style
// (`dgm:styleDef`) and layout (`dgm:layoutDef`, header and root algorithm only).
import { parseXml, type XmlDocument } from '../xml/index.js';
import { parseDrawingColorIn, parseDrawingColorList } from './drawing-color.js';
import {
	NS,
	children,
	descendants,
	elements,
	first,
	stringAttribute,
	type XmlElement,
} from './dom.js';
import { resolveDiagramLayoutCategory } from './layout-category.js';
import type {
	DiagramColorList,
	DiagramColorStyleLabel,
	DiagramColorsDefinition,
	DiagramDefinitionHeader,
	DiagramLayoutSummary,
	DiagramQuickStyleDefinition,
	DiagramStyleLabel,
	DiagramStyleRef,
} from './types.js';

const rootOf = (source: string | XmlDocument): XmlElement =>
	(typeof source === 'string' ? parseXml(source, { label: 'DiagramML' }) : source).documentElement;

/** `uniqueId`, `dgm:title`, `dgm:desc` and `dgm:catLst` shared by the definition parts. */
export function parseDefinitionHeader(root: XmlElement): DiagramDefinitionHeader {
	const header: DiagramDefinitionHeader = {
		uniqueId: stringAttribute(root, 'uniqueId') ?? '',
		categories: children(first(root, 'catLst', NS.dgm) ?? root, 'cat', NS.dgm).flatMap((cat) => {
			const type = stringAttribute(cat, 'type');
			const priority = Number.parseInt(cat.getAttribute('pri') ?? '', 10);
			return type ? [{ type, ...(Number.isFinite(priority) ? { priority } : {}) }] : [];
		}),
	};
	// An empty `val=""` title or description is the norm and carries no information.
	const title = stringAttribute(first(root, 'title', NS.dgm), 'val');
	const description = stringAttribute(first(root, 'desc', NS.dgm), 'val');
	if (title) header.title = title;
	if (description) header.description = description;
	return header;
}

function colorList(container: XmlElement | undefined): DiagramColorList {
	const method = stringAttribute(container, 'meth');
	return { ...(method ? { method } : {}), colors: parseDrawingColorList(container) };
}

/** Parses a colours definition part. */
export function parseDiagramColors(source: string | XmlDocument): DiagramColorsDefinition {
	const root = rootOf(source);
	const labels: DiagramColorStyleLabel[] = children(root, 'styleLbl', NS.dgm).map((label) => ({
		name: stringAttribute(label, 'name') ?? '',
		fill: colorList(first(label, 'fillClrLst', NS.dgm)),
		line: colorList(first(label, 'linClrLst', NS.dgm)),
		effect: colorList(first(label, 'effectClrLst', NS.dgm)),
		textLine: colorList(first(label, 'txLinClrLst', NS.dgm)),
		textFill: colorList(first(label, 'txFillClrLst', NS.dgm)),
		textEffect: colorList(first(label, 'txEffectClrLst', NS.dgm)),
	}));
	return { ...parseDefinitionHeader(root), labels };
}

function styleRef(element: XmlElement | undefined): DiagramStyleRef | undefined {
	if (!element) return undefined;
	const raw = element.getAttribute('idx');
	const ref: DiagramStyleRef = {};
	if (element.localName === 'fontRef') {
		if (raw) ref.fontIndex = raw;
	} else {
		const index = Number.parseInt(raw ?? '', 10);
		if (Number.isFinite(index)) ref.index = index;
	}
	const color = parseDrawingColorIn(element);
	if (color) ref.color = color;
	return ref;
}

/** The four style-matrix references of a `dgm:style` / `dsp:style` element. */
export function parseStyleReferences(
	style: XmlElement | undefined,
): Pick<DiagramStyleLabel, 'line' | 'fill' | 'effect' | 'font'> {
	const refs: Pick<DiagramStyleLabel, 'line' | 'fill' | 'effect' | 'font'> = {};
	const line = styleRef(first(style, 'lnRef', NS.a));
	const fill = styleRef(first(style, 'fillRef', NS.a));
	const effect = styleRef(first(style, 'effectRef', NS.a));
	const font = styleRef(first(style, 'fontRef', NS.a));
	if (line) refs.line = line;
	if (fill) refs.fill = fill;
	if (effect) refs.effect = effect;
	if (font) refs.font = font;
	return refs;
}

const hasContent = (element: XmlElement | undefined): boolean =>
	Boolean(element && (element.attributes.length > 0 || elements(element).length > 0));

/** Parses a quick style definition part. */
export function parseDiagramQuickStyle(source: string | XmlDocument): DiagramQuickStyleDefinition {
	const root = rootOf(source);
	const labels: DiagramStyleLabel[] = children(root, 'styleLbl', NS.dgm).map((label) => ({
		name: stringAttribute(label, 'name') ?? '',
		...parseStyleReferences(first(label, 'style', NS.dgm)),
		hasScene3d: Boolean(first(label, 'scene3d', NS.dgm)),
		hasShape3d: hasContent(first(label, 'sp3d', NS.dgm)),
	}));
	return {
		...parseDefinitionHeader(root),
		labels,
		has3d: labels.some((label) => label.hasShape3d),
	};
}

/** Parses the header, family and root algorithm of a layout definition part. */
export function parseDiagramLayoutSummary(source: string | XmlDocument): DiagramLayoutSummary {
	const root = rootOf(source);
	const header = parseDefinitionHeader(root);
	const family = resolveDiagramLayoutCategory(
		header.uniqueId,
		header.categories.map((category) => category.type),
	);
	const rootNode = first(root, 'layoutNode', NS.dgm);
	const rootAlgorithm = stringAttribute(first(rootNode, 'alg', NS.dgm), 'type');
	return {
		...header,
		...(family ? { family } : {}),
		...(rootAlgorithm ? { rootAlgorithm } : {}),
		layoutNodeCount: descendants(root, 'layoutNode').length,
	};
}
