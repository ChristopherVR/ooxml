import {
	NS,
	buildXml,
	children,
	elements,
	first,
	parseXml,
	type XmlElement,
} from '../../xml/index';
import type { DiagramColorScheme, DiagramStyleIntensity } from '../types';
import type { DiagramLayoutSource } from '../layout-preview';
import { buildFabricatedColorsXml, buildFabricatedQuickStyleXml } from '../write/fabricated-styles';
import { DIAGRAM_COLOR_ENTRIES, DIAGRAM_STYLE_ENTRIES } from '../catalogue';

export type DiagramFormatChange =
	| { kind: 'colors'; value: DiagramColorScheme }
	| { kind: 'style'; value: DiagramStyleIntensity };
const COLOR_NAMES = new Set(['schemeClr', 'srgbClr', 'scrgbClr', 'hslClr', 'prstClr', 'sysClr']);
function ensure(parent: XmlElement, name: string, namespace: string = NS.dgm): XmlElement {
	let node = first(parent, name, namespace);
	if (!node) {
		node = parent.ownerDocument.createElementNS(
			namespace,
			`${namespace === NS.a ? 'a' : 'dgm'}:${name}`,
		);
		const order =
			parent.localName === 'style'
				? ['lnRef', 'fillRef', 'effectRef', 'fontRef']
				: parent.localName === 'styleLbl'
					? ['scene3d', 'sp3d', 'txPr', 'style', 'extLst']
					: [];
		parent.insertBefore(
			node,
			elements(parent).find((child) => order.indexOf(child.localName) > order.indexOf(name)) ??
				null,
		);
	}
	return node;
}

/** Change supported fields in place, leaving extensions, foreign labels and other fields intact. */
export function writeDiagramFormatDefinition(
	source: string | undefined,
	change: DiagramFormatChange,
): string {
	if (
		!(change.kind === 'colors' ? DIAGRAM_COLOR_ENTRIES : DIAGRAM_STYLE_ENTRIES).some(
			(entry) => entry.id === change.value,
		)
	)
		throw new RangeError('Unknown SmartArt format preset.');
	const template =
		change.kind === 'colors'
			? buildFabricatedColorsXml(change.value)
			: buildFabricatedQuickStyleXml(change.value);
	const fresh = parseXml(template).documentElement;
	if (!source) {
		fresh.setAttribute('uniqueId', `urn:ooxml-core/${change.kind}/${change.value}`);
		return buildXml(fresh);
	}
	const doc = parseXml(source, { label: 'SmartArt formatting definition' });
	const root = doc.documentElement;
	if (
		root.namespaceURI !== NS.dgm ||
		root.localName !== (change.kind === 'colors' ? 'colorsDef' : 'styleDef')
	)
		throw new Error('The SmartArt formatting definition is invalid.');
	root.setAttribute('uniqueId', `urn:ooxml-core/${change.kind}/${change.value}`);
	for (const label of children(root, 'styleLbl', NS.dgm)) {
		const name = label.getAttribute('name') ?? '';
		const model =
			children(fresh, 'styleLbl', NS.dgm).find((entry) => entry.getAttribute('name') === name) ??
			children(fresh, 'styleLbl', NS.dgm)[0]!;
		if (change.kind === 'colors') {
			for (const listName of ['fillClrLst', 'linClrLst', 'txFillClrLst']) {
				const list = ensure(label, listName);
				const desired = first(model, listName, NS.dgm)!;
				list.setAttribute('meth', desired.getAttribute('meth') ?? 'repeat');
				for (const child of elements(list))
					if (child.namespaceURI === NS.a && COLOR_NAMES.has(child.localName))
						list.removeChild(child);
				for (const color of elements(desired)) list.appendChild(doc.importNode(color, true));
			}
		} else {
			const style = ensure(label, 'style');
			const desired = first(model, 'style', NS.dgm)!;
			for (const key of ['lnRef', 'fillRef', 'effectRef']) {
				const ref = ensure(style, key, NS.a);
				const templateRef = first(desired, key, NS.a)!;
				ref.setAttribute('idx', templateRef.getAttribute('idx') ?? '0');
				if (
					!elements(ref).some(
						(child) => child.namespaceURI === NS.a && COLOR_NAMES.has(child.localName),
					)
				)
					for (const child of elements(templateRef)) ref.appendChild(doc.importNode(child, true));
			}
			if (!first(style, 'fontRef', NS.a)) {
				const ref = ensure(style, 'fontRef', NS.a);
				ref.setAttribute('idx', 'minor');
			}
		}
	}
	return buildXml(doc);
}

export function formatDiagramSource(
	source: DiagramLayoutSource | undefined,
	change: DiagramFormatChange,
): {
	source: DiagramLayoutSource;
	colorsId?: string;
	quickStyleId?: string;
} {
	const xml = writeDiagramFormatDefinition(
		change.kind === 'colors' ? source?.colorsXml : source?.quickStyleXml,
		change,
	);
	const id = `urn:ooxml-core/${change.kind}/${change.value}`;
	return change.kind === 'colors'
		? { source: { ...source, colorsXml: xml }, colorsId: id }
		: { source: { ...source, quickStyleXml: xml }, quickStyleId: id };
}
