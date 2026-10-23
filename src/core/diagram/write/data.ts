/**
 * Fabricate a diagram DATA part (`dgm:dataModel`, `ppt/diagrams/dataN.xml`,
 * `word/diagrams/dataN.xml`, `xl/diagrams/dataN.xml`) for a SmartArt graphic
 * created in memory, with no existing diagram parts in the package. Moved from
 * `pptx/core/core/runtime/smartart-fabrication-data.ts` and made generic over
 * the neutral node model; the model id generator is injected.
 *
 * The fabricated data model is deliberately minimal: a `doc` point, one
 * content point per in-memory node (plus the `parTrans` / `sibTrans` support
 * pair every real content point carries), and the `parOf` connections that
 * wire the tree. No `pres`-type presentation points are emitted; PowerPoint
 * computes the presentation tree itself from the layout definition on open.
 *
 * Model ids MUST be `{GUID}` strings: schema-valid informal ids are still
 * rejected by PowerPoint's loader as a corrupt file (confirmed empirically
 * via COM automation). In-memory node ids produced by the editing UI are
 * informal (`node-<timestamp>-…`), so every id is remapped here.
 */
import { createOfficeGuid } from '../../crypto/uuid';
import type { DiagramNode } from '../model/node';

export const DIAGRAM_XML_PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const DIAGRAM_DATA_XMLNS =
	'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** Escape a text value for inclusion in XML content or attributes. */
export function escapeDiagramXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&apos;');
}

/**
 * `a:avLst` for a fabricated preset geometry. Keeping the cached adjustments
 * matters: a Basic Pyramid tier is a `trapezoid` at `adj 95238`, and an empty
 * list would redraw it at the 25000 default.
 */
export function diagramAvLstXml(adjustments: Record<string, number> | undefined): string {
	const entries = Object.entries(adjustments ?? {}).filter(([, v]) => Number.isFinite(v));
	if (entries.length === 0) {
		return '<a:avLst/>';
	}
	const gds = entries
		.map(([name, v]) => `<a:gd name="${escapeDiagramXml(name)}" fmla="val ${Math.round(v)}"/>`)
		.join('');
	return `<a:avLst>${gds}</a:avLst>`;
}

const GUID_MODEL_ID = /^\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$/u;

/** Unique-id URNs of the fabricated layout / quick-style / colors parts. */
export interface DiagramDefinitionIds {
	layoutUniqueId: string;
	layoutCategory: string;
	quickStyleUniqueId: string;
	colorsUniqueId: string;
	colorsCategory: string;
}

/** The empty `dgm:t` body used for doc / parTrans / sibTrans points. */
const EMPTY_TEXT_BODY =
	'<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>';

function contentTextBody(text: string): string {
	if (!text) {
		return EMPTY_TEXT_BODY;
	}
	return `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeDiagramXml(text)}</a:t></a:r></a:p></dgm:t>`;
}

function contentPointXml(guid: string, node: DiagramNode): string {
	const prSet = node.text ? '<dgm:prSet/>' : '<dgm:prSet phldrT="[Text]" phldr="1"/>';
	// Only `type="asst"` (ECMA-376 ST_DataPointType) is a valid content-point
	// override beyond the schema default ("node", expressed by omitting the
	// attribute); dropping it here reloaded an SDK-authored assistant node as
	// an ordinary child, losing its org-chart placement on every round-trip.
	const typeAttr = node.nodeType === 'asst' ? ' type="asst"' : '';
	return `<dgm:pt modelId="${guid}"${typeAttr}>${prSet}<dgm:spPr/>${contentTextBody(node.text ?? '')}</dgm:pt>`;
}

function transPointXml(guid: string, type: 'parTrans' | 'sibTrans', cxnGuid: string): string {
	return `<dgm:pt modelId="${guid}" type="${type}" cxnId="${cxnGuid}"><dgm:prSet/><dgm:spPr/>${EMPTY_TEXT_BODY}</dgm:pt>`;
}

function docPointXml(docGuid: string, ids: DiagramDefinitionIds): string {
	return (
		`<dgm:pt modelId="${docGuid}" type="doc">` +
		`<dgm:prSet loTypeId="${escapeDiagramXml(ids.layoutUniqueId)}" loCatId="${escapeDiagramXml(ids.layoutCategory)}"` +
		` qsTypeId="${escapeDiagramXml(ids.quickStyleUniqueId)}" qsCatId="simple"` +
		` csTypeId="${escapeDiagramXml(ids.colorsUniqueId)}" csCatId="${escapeDiagramXml(ids.colorsCategory)}"/>` +
		`<dgm:spPr/>${EMPTY_TEXT_BODY}</dgm:pt>`
	);
}

/**
 * Build the stable `nodeId -> {GUID}` model-id map for a fabricated diagram.
 *
 * Shared between the data part (point model ids) and the cached drawing part
 * (`dsp:sp` model ids) so PowerPoint correlates each drawn shape with its
 * data point and keeps the cached geometry instead of recomputing it.
 */
export function buildDiagramNodeGuidMap(
	nodes: readonly DiagramNode[],
	createModelId: () => string = createOfficeGuid,
): Map<string, string> {
	const guidByNodeId = new Map<string, string>();
	for (const node of nodes) {
		if (node.id && !guidByNodeId.has(node.id)) {
			guidByNodeId.set(node.id, GUID_MODEL_ID.test(node.id) ? node.id : createModelId());
		}
	}
	return guidByNodeId;
}

/** Build stable presentation-point ids for cached drawing shapes. */
export function buildDiagramPresentationGuidMap(
	nodes: readonly DiagramNode[],
	createModelId: () => string = createOfficeGuid,
): Map<string, string> {
	const guidByNodeId = new Map<string, string>();
	for (const node of nodes) {
		if (node.id && !guidByNodeId.has(node.id)) {
			guidByNodeId.set(node.id, createModelId());
		}
	}
	return guidByNodeId;
}

/** Options controlling optional data-model extensions. */
export interface FabricatedDiagramDataOptions {
	/**
	 * Relationship id in the owning slide's rels file for the cached
	 * `drawingN.xml`. When set, a `dsp:dataModelExt` extension is emitted so
	 * PowerPoint loads the cached drawing instead of recomputing the layout.
	 */
	drawingRelId?: string;
	/**
	 * Pre-built `nodeId -> {GUID}` map. Supply the same map used to build the
	 * cached drawing so point and shape model ids match. Built internally when
	 * omitted.
	 */
	guidByNodeId?: Map<string, string>;
	/** Presentation-point ids used by cached `dsp:sp` shapes. */
	presentationGuidByNodeId?: Map<string, string>;
	/** Mints point and connection model ids; a braced upper-case GUID by default. */
	createModelId?: () => string;
}

const DSP_DATA_MODEL_EXT_URI = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

/**
 * Build the complete `dataN.xml` payload for a fabricated SmartArt diagram.
 *
 * Node parent wiring comes from each node's `parentId`; a missing or
 * unresolvable `parentId` parents the node under the `doc` point (a
 * top-level entry). `srcOrd` is assigned per parent in `nodes` order so
 * sibling order in the viewer is preserved in the saved diagram.
 */
export function writeFabricatedDiagramData(
	data: { nodes: readonly DiagramNode[] },
	ids: DiagramDefinitionIds,
	options: FabricatedDiagramDataOptions = {},
): string {
	const createModelId = options.createModelId ?? createOfficeGuid;
	const docGuid = createModelId();
	const guidByNodeId = options.guidByNodeId ?? buildDiagramNodeGuidMap(data.nodes, createModelId);
	const presentationGuidByNodeId =
		options.presentationGuidByNodeId ?? buildDiagramPresentationGuidMap(data.nodes, createModelId);

	const points: string[] = [docPointXml(docGuid, ids)];
	const connections: string[] = [];
	const nextSrcOrd = new Map<string, number>();

	for (const node of data.nodes) {
		const guid = node.id ? guidByNodeId.get(node.id) : undefined;
		if (!guid) {
			continue;
		}
		const parentGuid = (node.parentId && guidByNodeId.get(node.parentId)) || docGuid;
		const srcOrd = nextSrcOrd.get(parentGuid) ?? 0;
		nextSrcOrd.set(parentGuid, srcOrd + 1);

		const cxnGuid = createModelId();
		const parTransGuid = createModelId();
		const sibTransGuid = createModelId();

		points.push(contentPointXml(guid, node));
		points.push(transPointXml(parTransGuid, 'parTrans', cxnGuid));
		points.push(transPointXml(sibTransGuid, 'sibTrans', cxnGuid));
		connections.push(
			`<dgm:cxn modelId="${cxnGuid}" srcId="${parentGuid}" destId="${guid}" srcOrd="${srcOrd}" destOrd="0" parTransId="${parTransGuid}" sibTransId="${sibTransGuid}"/>`,
		);

		const presentationGuid = node.id ? presentationGuidByNodeId.get(node.id) : undefined;
		if (presentationGuid) {
			points.push(
				`<dgm:pt modelId="${presentationGuid}" type="pres">` +
					`<dgm:prSet presAssocID="${guid}" presName="node" presStyleLbl="node1"` +
					` presStyleIdx="${srcOrd}" presStyleCnt="${data.nodes.length}"/>` +
					`<dgm:spPr/></dgm:pt>`,
			);
			connections.push(
				`<dgm:cxn modelId="${createModelId()}" type="presOf" srcId="${guid}"` +
					` destId="${presentationGuid}" srcOrd="0" destOrd="0"` +
					` presId="${escapeDiagramXml(ids.layoutUniqueId)}"/>`,
			);
		}
	}

	return (
		`${DIAGRAM_XML_PROLOG}\r\n<dgm:dataModel ${DIAGRAM_DATA_XMLNS}>` +
		`<dgm:ptLst>${points.join('')}</dgm:ptLst>` +
		`<dgm:cxnLst>${connections.join('')}</dgm:cxnLst>` +
		`<dgm:bg/><dgm:whole/>${dataModelExtXml(options.drawingRelId)}</dgm:dataModel>`
	);
}

/**
 * Build the `dgm:extLst` carrying the `dsp:dataModelExt` that links the data
 * model to its cached drawing part. Empty when no drawing rel is supplied.
 */
function dataModelExtXml(drawingRelId: string | undefined): string {
	if (!drawingRelId) {
		return '';
	}
	return (
		`<dgm:extLst>` +
		`<a:ext uri="${DSP_DATA_MODEL_EXT_URI}">` +
		`<dsp:dataModelExt xmlns:dsp="${DSP_DATA_MODEL_EXT_URI}" relId="${escapeDiagramXml(drawingRelId)}"` +
		` minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>` +
		`</a:ext>` +
		`</dgm:extLst>`
	);
}

/** `xml` with exactly one XML declaration (the built-in definitions carry one; others may not). */
export function withDiagramXmlProlog(xml: string): string {
	return /^\s*<\?xml\b/u.test(xml) ? xml : `${DIAGRAM_XML_PROLOG}\r\n${xml}`;
}
