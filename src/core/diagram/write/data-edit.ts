/**
 * Write an edited node tree back into an EXISTING diagram data part (`dgm:dataModel`), on the
 * shared DOM. Everything the edit does not touch keeps its markup: unedited points, their
 * `dgm:spPr` and `dgm:prSet` (custom layout, placeholders), presentation points of nodes that
 * remain, the background, `dgm:whole` and unknown extensions. Only what the edit changes is
 * rewritten:
 *
 * - a removed node loses its point, its `parOf` connection with the transition points that go with
 *   it, and the presentation points (and their `presOf`/`presParOf` connections) laid out for them;
 * - a new node gets a content point, its `parTrans`/`sibTrans` pair and a `parOf` connection;
 * - every remaining node's `parOf` connection is pointed at its (possibly new) parent with
 *   `srcOrd` from its position among its siblings in `nodes`;
 * - a node whose text changed gets a new text body, keeping the first paragraph's `a:pPr` and the
 *   first run's `a:rPr` (later runs' formatting is not kept for edited text).
 *
 * Presentation points of new nodes are not fabricated: Word lays a diagram out again from the data
 * and layout parts and rebuilds them (evidence in docs/agnostic-core-plan.md, "Word and an edited
 * data part"). `dropDrawingLink` removes the `dsp:dataModelExt` that ties the data to a cached
 * drawing, for a host that drops the stale drawing part.
 */
import { createOfficeGuid } from '../../crypto/uuid';
import {
	NS,
	buildXml,
	elements,
	parseXml,
	type XmlDocument,
	type XmlElement,
} from '../../xml/index';

const DGM = NS.dgm;
const A = NS.a;
const DSP_EXT_URI = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const GUID_MODEL_ID = /^\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$/u;
const CONTENT_TYPES = new Set(['', 'node', 'asst']);

/** A content node as the writer sees it: array order is sibling order. */
export interface DiagramDataEditNode {
	id: string;
	text: string;
	parentId?: string | undefined;
	/** `asst` keeps an assistant point's type on a new point. */
	nodeType?: string | undefined;
}

export interface DiagramDataEditOptions {
	/** Mints the model ids of new points and connections; a braced upper-case GUID by default. */
	createModelId?: () => string;
	/** Remove the `dsp:dataModelExt` link to the cached drawing. */
	dropDrawingLink?: boolean;
}

const child = (parent: XmlElement, local: string, ns: string = DGM): XmlElement | undefined =>
	elements(parent).find((el) => el.localName === local && el.namespaceURI === ns);

const pointType = (pt: XmlElement): string => pt.getAttribute('type') ?? '';
const isParOf = (cxn: XmlElement): boolean => (cxn.getAttribute('type') ?? 'parOf') === 'parOf';

/** The text of a `dgm:t`: its paragraphs joined by a newline (as `parseDiagramDataModel` reads it). */
function bodyText(pt: XmlElement): string {
	const t = child(pt, 't');
	if (!t) return '';
	return elements(t)
		.filter((el) => el.localName === 'p' && el.namespaceURI === A)
		.map((p) =>
			elements(p)
				.map((item) => {
					if (item.localName === 'br') return '\n';
					return item.localName === 'r' || item.localName === 'fld'
						? (child(item, 't', A)?.textContent ?? '')
						: '';
				})
				.join(''),
		)
		.join('\n');
}

function element(doc: XmlDocument, ns: string, qualified: string): XmlElement {
	return doc.createElementNS(ns, qualified);
}

/** Replace the paragraphs of a point's `dgm:t` with `text`, keeping the first pPr and rPr. */
function writeBodyText(doc: XmlDocument, pt: XmlElement, text: string): void {
	let t = child(pt, 't');
	if (!t) {
		t = element(doc, DGM, 'dgm:t');
		t.appendChild(element(doc, A, 'a:bodyPr'));
		t.appendChild(element(doc, A, 'a:lstStyle'));
		pt.appendChild(t);
	}
	const paragraphs = elements(t).filter((el) => el.localName === 'p' && el.namespaceURI === A);
	const firstParagraph = paragraphs[0];
	const pPr = firstParagraph && child(firstParagraph, 'pPr', A);
	const firstRun = paragraphs
		.flatMap((p) => elements(p))
		.find((el) => el.localName === 'r' && el.namespaceURI === A);
	const rPr =
		(firstRun && child(firstRun, 'rPr', A)) ??
		(firstParagraph && child(firstParagraph, 'endParaRPr', A));
	const anchor = paragraphs.at(-1)?.nextSibling ?? null;
	for (const p of paragraphs) t.removeChild(p);
	const runProperties = (local: 'rPr' | 'endParaRPr'): XmlElement => {
		const next = element(doc, A, `a:${local}`);
		for (const attribute of Array.from(rPr?.attributes ?? []))
			next.setAttribute(attribute.name, attribute.value);
		if (!next.hasAttribute('lang')) next.setAttribute('lang', 'en-US');
		for (const node of Array.from(rPr?.childNodes ?? [])) next.appendChild(node.cloneNode(true));
		return next;
	};
	for (const line of text.split('\n')) {
		const p = element(doc, A, 'a:p');
		if (pPr) p.appendChild(pPr.cloneNode(true));
		if (line) {
			const r = element(doc, A, 'a:r');
			r.appendChild(runProperties('rPr'));
			const value = element(doc, A, 'a:t');
			value.appendChild(doc.createTextNode(line));
			r.appendChild(value);
			p.appendChild(r);
		} else p.appendChild(runProperties('endParaRPr'));
		t.insertBefore(p, anchor);
	}
	const prSet = child(pt, 'prSet');
	if (text && prSet?.getAttribute('phldr') === '1') prSet.removeAttribute('phldr');
}

function emptyBody(doc: XmlDocument): XmlElement {
	const t = element(doc, DGM, 'dgm:t');
	t.appendChild(element(doc, A, 'a:bodyPr'));
	t.appendChild(element(doc, A, 'a:lstStyle'));
	const p = element(doc, A, 'a:p');
	const end = element(doc, A, 'a:endParaRPr');
	end.setAttribute('lang', 'en-US');
	p.appendChild(end);
	t.appendChild(p);
	return t;
}

function newPoint(
	doc: XmlDocument,
	modelId: string,
	type?: string,
	cxnId?: string,
	placeholder = false,
): XmlElement {
	const pt = element(doc, DGM, 'dgm:pt');
	pt.setAttribute('modelId', modelId);
	if (type) pt.setAttribute('type', type);
	if (cxnId) pt.setAttribute('cxnId', cxnId);
	const prSet = element(doc, DGM, 'dgm:prSet');
	if (placeholder) prSet.setAttribute('phldrT', '[Text]');
	pt.appendChild(prSet);
	pt.appendChild(element(doc, DGM, 'dgm:spPr'));
	pt.appendChild(emptyBody(doc));
	return pt;
}

/** Remove the `a:ext` holding `dsp:dataModelExt` (and an `dgm:extLst` it leaves empty). */
function dropDrawingExtension(model: XmlElement): void {
	const extLst = child(model, 'extLst');
	if (!extLst) return;
	for (const ext of elements(extLst))
		if (ext.getAttribute('uri') === DSP_EXT_URI) extLst.removeChild(ext);
	if (elements(extLst).length === 0) model.removeChild(extLst);
}

/**
 * The data part `xml` with its content points replaced by `nodes` (tree order: a node's position
 * among the nodes with the same parent is its sibling order). A node whose `parentId` is not one
 * of `nodes` hangs under the document point. Throws when `xml` has no document point.
 */
export function writeEditedDiagramData(
	xml: string,
	nodes: readonly DiagramDataEditNode[],
	options: DiagramDataEditOptions = {},
): string {
	const createModelId = options.createModelId ?? createOfficeGuid;
	const doc = parseXml(xml, { label: 'diagram data' });
	const model = doc.documentElement as XmlElement;
	const ptLst = child(model, 'ptLst');
	if (!ptLst) throw new Error('The diagram data part has no point list.');
	let cxnLst = child(model, 'cxnLst');
	if (!cxnLst) {
		cxnLst = element(doc, DGM, 'dgm:cxnLst');
		ptLst.parentNode?.insertBefore(cxnLst, ptLst.nextSibling);
	}
	const points = elements(ptLst).filter((el) => el.localName === 'pt');
	const pointById = new Map(points.map((pt) => [pt.getAttribute('modelId') ?? '', pt]));
	const docPoint = points.find((pt) => pointType(pt) === 'doc');
	const rootId = docPoint?.getAttribute('modelId');
	if (!rootId) throw new Error('The diagram data part has no document point.');
	const connections = () => elements(cxnLst).filter((el) => el.localName === 'cxn');

	// New ids that are not GUIDs (an editor's informal ids) are minted, and parents follow.
	const idMap = new Map<string, string>();
	for (const node of nodes) {
		const existing = pointById.get(node.id);
		const valid = existing ? CONTENT_TYPES.has(pointType(existing)) : GUID_MODEL_ID.test(node.id);
		idMap.set(node.id, valid ? node.id : createModelId());
	}
	const kept = new Set(idMap.values());

	// Removed content points: the point, its parOf edges and transitions, and the presentation
	// points laid out for any of them.
	const removed = new Set<string>();
	for (const pt of points) {
		const id = pt.getAttribute('modelId') ?? '';
		if (CONTENT_TYPES.has(pointType(pt)) && !kept.has(id)) removed.add(id);
	}
	const parOfByDest = new Map<string, XmlElement>();
	for (const cxn of connections()) {
		if (!isParOf(cxn)) continue;
		const dest = cxn.getAttribute('destId') ?? '';
		if (removed.has(dest) || removed.has(cxn.getAttribute('srcId') ?? '')) {
			for (const key of ['parTransId', 'sibTransId']) {
				const id = cxn.getAttribute(key);
				if (id) removed.add(id);
			}
			const cxnId = cxn.getAttribute('modelId');
			for (const pt of points)
				if (cxnId && pt.getAttribute('cxnId') === cxnId)
					removed.add(pt.getAttribute('modelId') ?? '');
			cxnLst.removeChild(cxn);
		} else if (!parOfByDest.has(dest)) parOfByDest.set(dest, cxn);
	}
	const removedPres = new Set<string>();
	for (const pt of points)
		if (
			pointType(pt) === 'pres' &&
			removed.has(child(pt, 'prSet')?.getAttribute('presAssocID') ?? '')
		)
			removedPres.add(pt.getAttribute('modelId') ?? '');
	for (let grew = true; grew;) {
		grew = false;
		for (const cxn of connections())
			if (
				cxn.getAttribute('type') === 'presParOf' &&
				removedPres.has(cxn.getAttribute('srcId') ?? '') &&
				!removedPres.has(cxn.getAttribute('destId') ?? '')
			) {
				removedPres.add(cxn.getAttribute('destId') ?? '');
				grew = true;
			}
	}
	for (const cxn of connections()) {
		const type = cxn.getAttribute('type');
		const src = cxn.getAttribute('srcId') ?? '';
		const dest = cxn.getAttribute('destId') ?? '';
		if (
			(type === 'presOf' && (removed.has(src) || removedPres.has(dest))) ||
			(type === 'presParOf' && (removedPres.has(src) || removedPres.has(dest)))
		)
			cxnLst.removeChild(cxn);
	}
	for (const pt of points) {
		const id = pt.getAttribute('modelId') ?? '';
		if (removed.has(id) || removedPres.has(id)) ptLst.removeChild(pt);
	}

	// Remaining and new nodes: text, parent and sibling order.
	const nextOrd = new Map<string, number>();
	for (const node of nodes) {
		const id = idMap.get(node.id) as string;
		const parentId = (node.parentId && idMap.get(node.parentId)) || rootId;
		const srcOrd = nextOrd.get(parentId) ?? 0;
		nextOrd.set(parentId, srcOrd + 1);
		let pt = pointById.get(id);
		if (pt && !removed.has(id)) {
			if (bodyText(pt) !== node.text) writeBodyText(doc, pt, node.text);
		} else {
			pt = newPoint(doc, id, node.nodeType === 'asst' ? 'asst' : undefined, undefined, true);
			if (node.text) writeBodyText(doc, pt, node.text);
			ptLst.appendChild(pt);
		}
		const cxn = parOfByDest.get(id);
		if (cxn) {
			cxn.setAttribute('srcId', parentId);
			cxn.setAttribute('srcOrd', String(srcOrd));
			continue;
		}
		const cxnId = createModelId();
		const parTransId = createModelId();
		const sibTransId = createModelId();
		ptLst.appendChild(newPoint(doc, parTransId, 'parTrans', cxnId));
		ptLst.appendChild(newPoint(doc, sibTransId, 'sibTrans', cxnId));
		const edge = element(doc, DGM, 'dgm:cxn');
		edge.setAttribute('modelId', cxnId);
		edge.setAttribute('srcId', parentId);
		edge.setAttribute('destId', id);
		edge.setAttribute('srcOrd', String(srcOrd));
		edge.setAttribute('destOrd', '0');
		edge.setAttribute('parTransId', parTransId);
		edge.setAttribute('sibTransId', sibTransId);
		cxnLst.appendChild(edge);
	}
	if (options.dropDrawingLink) dropDrawingExtension(model);
	return buildXml(doc);
}
