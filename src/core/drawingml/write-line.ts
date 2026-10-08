// DrawingML outline writer (`a:ln`, ECMA-376 Part 1, 20.1.2.2.24): the modelled width, cap, fill
// and preset dash, written fresh or patched into an existing outline so its join, compound type,
// arrow heads and extensions survive.
import { NS, buildXml, elements, parseXml, type XmlElement } from '../xml/index';
import { drawingFillXml } from './write-fill';
import type { DrawingLine } from './types';

const FILLS = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill']);

/** Removes the `a` declarations the colour writer repeats inside an imported fill. */
function dropDrawingDeclarations(element: XmlElement): void {
	if (element.getAttribute('xmlns:a') === NS.a) element.removeAttribute('xmlns:a');
	for (const child of elements(element)) dropDrawingDeclarations(child);
}

/** Applies `line` to an `a:ln` element; children the model does not cover are kept. */
export function patchDrawingLine(ln: XmlElement, line: DrawingLine): void {
	if (line.widthEmu === undefined) ln.removeAttribute('w');
	else ln.setAttribute('w', String(Math.round(line.widthEmu)));
	if (line.cap === undefined) ln.removeAttribute('cap');
	else ln.setAttribute('cap', line.cap);
	// A fill the shared writer cannot express (pattern, picture) stays as written.
	const fill = line.fill && drawingFillXml(line.fill);
	if (fill || !line.fill)
		for (const node of elements(ln))
			if (node.namespaceURI === NS.a && FILLS.has(node.localName)) ln.removeChild(node);
	if (fill) {
		const parsed = parseXml(`<w xmlns:a="${NS.a}">${fill}</w>`).documentElement;
		for (const node of elements(parsed).reverse()) {
			const imported = ln.ownerDocument.importNode(node, true);
			dropDrawingDeclarations(imported);
			ln.insertBefore(imported, ln.firstChild);
		}
	}
	const dashes = elements(ln).filter(
		(node) =>
			node.namespaceURI === NS.a &&
			(node.localName === 'prstDash' || node.localName === 'custDash'),
	);
	if (line.dash === undefined) {
		for (const node of dashes) if (node.localName === 'prstDash') ln.removeChild(node);
		return;
	}
	const dash = ln.ownerDocument.createElementNS(NS.a, 'a:prstDash');
	dash.setAttribute('val', line.dash);
	for (const node of dashes) ln.removeChild(node);
	const afterFill = elements(ln).find(
		(node) => node.namespaceURI !== NS.a || !FILLS.has(node.localName),
	);
	ln.insertBefore(dash, afterFill ?? null);
}

/** A new `a:ln` for `line`. Fills the shared fill writer cannot express are left out. */
export function drawingLineXml(line: DrawingLine): string {
	const doc = parseXml(`<a:ln xmlns:a="${NS.a}"/>`);
	patchDrawingLine(doc.documentElement, line);
	return buildXml(doc);
}
