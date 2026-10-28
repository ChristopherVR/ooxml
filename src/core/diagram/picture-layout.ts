import { NS, buildXml, children, first, parseXml, type XmlElement } from '../xml/index';
import { parseDrawingPictureFill } from '../drawingml/picture-fill';
import type { DrawingFill } from '../drawingml/types';
import type { DiagramDataModel, DiagramDrawing, DiagramIssue } from './types';
import type { DiagramLayoutSource } from './layout-preview';
import type { RenderedNode } from './layout/smartart-layout-types';
import {
	parseRelationships,
	buildRelationshipsXml,
	resolvePartPath,
	RELATIONSHIP_TYPES,
} from '../opc/index';

export type DiagramPictureFill = Extract<DrawingFill, { kind: 'picture' }>;
export interface DiagramPictureSource {
	nodeId: string;
	role?: string;
	styleLabel?: string;
	fill: DiagramPictureFill;
}
export const DIAGRAM_PICTURE_SOURCE_NS = 'urn:ooxml-core:diagram-picture-sources:1';

function remapFill(fill: DiagramPictureFill, id: string): DiagramPictureFill {
	if (!fill.sourceXml) return { ...fill, relId: id };
	const node = parseXml(fill.sourceXml).documentElement;
	const blip = first(node, 'blip', NS.a);
	if (blip) {
		blip.setAttributeNS(NS.r, 'r:embed', id);
		blip.removeAttributeNS(NS.r, 'link');
	}
	return { ...fill, relId: id, sourceXml: buildXml(node) };
}

/** Retain content identities, never infer picture ownership from text or cached shape order. */
export function captureDiagramPictures(
	model: DiagramDataModel,
	dataXml: string,
	dataParts: Readonly<Record<string, string>>,
	drawing?: DiagramDrawing,
): Pick<DiagramLayoutSource, 'pictures' | 'pictureParts'> {
	const pictures: DiagramPictureSource[] = [];
	const parts = { ...dataParts };
	const add = (source: DiagramPictureSource, owner: Readonly<Record<string, string>>) => {
		const part = source.fill.relId ? owner[source.fill.relId] : undefined;
		// External or unresolved media remains in the original source, never promoted into a cache.
		if (!part) return;
		let id = Object.keys(parts).find((key) => parts[key] === part);
		if (!id) {
			let index = 1;
			while (Object.hasOwn(parts, `ooxmlDiagramPicture${index}`)) index++;
			id = `ooxmlDiagramPicture${index}`;
			parts[id] = part;
		}
		const entry = { ...source, fill: remapFill(source.fill, id) };
		const existing = pictures.findIndex(
			(item) =>
				item.nodeId === entry.nodeId &&
				((item.role === entry.role && item.styleLabel === entry.styleLabel) ||
					(!item.role && !item.styleLabel && item.fill.relId === entry.fill.relId)),
		);
		if (existing < 0) pictures.push(entry);
		else pictures[existing] = entry;
	};
	const root = parseXml(dataXml).documentElement;
	const saved = root.getElementsByTagNameNS(DIAGRAM_PICTURE_SOURCE_NS, 'picture');
	for (const element of saved) {
		const nodeId = element.getAttribute('nodeId');
		const fill = first(element, 'blipFill', NS.a);
		if (nodeId && fill && model.nodes.some((node) => node.modelId === nodeId))
			add(
				{
					nodeId,
					fill: parseDrawingPictureFill(fill),
					...(element.getAttribute('role') ? { role: element.getAttribute('role')! } : {}),
					...(element.getAttribute('styleLabel')
						? { styleLabel: element.getAttribute('styleLabel')! }
						: {}),
				},
				dataParts,
			);
	}
	for (const point of model.points) {
		if (point.fill?.kind !== 'picture') continue;
		const nodeId = point.type === 'pres' ? point.presAssocId : point.modelId;
		if (
			nodeId &&
			model.nodes.some((node) => node.modelId === nodeId) &&
			!pictures.some((item) => item.nodeId === nodeId)
		)
			add(
				{
					nodeId,
					fill: point.fill,
					...(point.presName ? { role: point.presName } : {}),
					...(point.presStyleLabel ? { styleLabel: point.presStyleLabel } : {}),
				},
				dataParts,
			);
	}
	for (const shape of drawing?.shapes ?? []) {
		if (shape.fill?.kind !== 'picture') continue;
		const point = model.points.find((point) => point.modelId === shape.modelId);
		const nodeId =
			shape.sourceNodeId ?? (point?.type === 'pres' ? point.presAssocId : point?.modelId);
		// A generated cache may use a different layout role after a layout switch. The saved
		// association remains authoritative; do not manufacture a second image for that node.
		const part = shape.fill.relId ? drawing?.pictureParts?.[shape.fill.relId] : undefined;
		if (
			saved.length &&
			pictures.some(
				(item) =>
					item.nodeId === nodeId &&
					item.fill.relId &&
					parts[item.fill.relId] === part &&
					item.role !== point?.presName &&
					item.styleLabel !== point?.presStyleLabel,
			)
		)
			continue;
		if (nodeId && model.nodes.some((node) => node.modelId === nodeId))
			add(
				{
					nodeId,
					fill: shape.fill,
					...(point?.presName ? { role: point.presName } : {}),
					...(point?.presStyleLabel ? { styleLabel: point.presStyleLabel } : {}),
				},
				drawing?.pictureParts ?? {},
			);
	}
	return pictures.length ? { pictures, pictureParts: parts } : {};
}

/** Image placeholders consume pictures; sibling text roles never acquire the image as a fill. */
export function applyDiagramPictureSources(
	drawing: DiagramDrawing,
	nodes: readonly RenderedNode[],
	source: DiagramLayoutSource | undefined,
	issues: DiagramIssue[],
	liveNodeIds?: ReadonlySet<string>,
): void {
	if (!source?.pictures?.length) return;
	if (source.pictureParts) drawing.pictureParts = { ...source.pictureParts };
	const used = new Set<DiagramPictureSource>();
	nodes.forEach((node, index) => {
		if (!node.picturePlaceholder || !node.nodeId) return;
		const shape = drawing.shapes[index];
		if (!shape) return;
		const candidates = source.pictures!.filter((picture) => picture.nodeId === node.nodeId);
		const matching = candidates.filter(
			(picture) =>
				picture.role === node.layoutRole ||
				(!!picture.styleLabel && picture.styleLabel === node.styleLabel),
		);
		const picture =
			matching.length === 1 ? matching[0] : candidates.length === 1 ? candidates[0] : undefined;
		if (picture) {
			shape.fill = structuredClone(picture.fill);
			delete shape.text;
			delete shape.body;
			used.add(picture);
		} else if (candidates.length)
			issues.push({
				code: 'DIAGRAM_PICTURE_AMBIGUOUS',
				message: `Picture role ${node.layoutRole ?? node.key} has multiple source pictures; no arbitrary assignment was made.`,
			});
	});
	if (
		source.pictures.some(
			(picture) => (!liveNodeIds || liveNodeIds.has(picture.nodeId)) && !used.has(picture),
		)
	)
		issues.push({
			code: 'DIAGRAM_PICTURE_LAYOUT_UNSUPPORTED',
			message:
				'Some retained pictures have no matching image placeholder in this layout. Their source associations are preserved for a compatible layout.',
		});
}

/** Persist associations even when changing to a layout without a compatible picture role. */
export function withDiagramPictureSources(
	dataXml: string,
	source: DiagramLayoutSource | undefined,
): string {
	if (!source?.pictures?.length) return dataXml;
	const doc = parseXml(dataXml);
	const root = doc.documentElement;
	const points = first(root, 'ptLst', NS.dgm);
	const live = children(points ?? root, 'pt', NS.dgm).filter(
		(point) =>
			!['doc', 'pres', 'parTrans', 'sibTrans'].includes(point.getAttribute('type') ?? 'node'),
	);
	for (const element of [...root.getElementsByTagNameNS(DIAGRAM_PICTURE_SOURCE_NS, 'pictures')]) {
		const extension = element.parentNode;
		if (extension?.parentNode) extension.parentNode.removeChild(extension);
	}
	const retained = source.pictures.filter((picture) =>
		live.some((point) => point.getAttribute('modelId') === picture.nodeId),
	);
	if (!retained.length) return buildXml(doc);
	const list = first(root, 'extLst', NS.dgm) ?? doc.createElementNS(NS.dgm, 'dgm:extLst');
	if (!list.parentNode) root.appendChild(list);
	const extension = doc.createElementNS(NS.a, 'a:ext');
	extension.setAttribute('uri', DIAGRAM_PICTURE_SOURCE_NS);
	const payload = doc.createElementNS(DIAGRAM_PICTURE_SOURCE_NS, 'op:pictures');
	extension.appendChild(payload);
	list.appendChild(extension);
	for (const picture of retained) {
		if (!picture.fill.sourceXml)
			throw new Error('A retained SmartArt picture needs source fill XML');
		const element = doc.createElementNS(DIAGRAM_PICTURE_SOURCE_NS, 'op:picture');
		element.setAttribute('nodeId', picture.nodeId);
		if (picture.role) element.setAttribute('role', picture.role);
		if (picture.styleLabel) element.setAttribute('styleLabel', picture.styleLabel);
		const fill = parseXml(picture.fill.sourceXml).documentElement;
		element.appendChild(doc.importNode(fill, true));
		payload.appendChild(element);
		const presentations = children(points ?? root, 'pt', NS.dgm).filter((point) => {
			const props = first(point, 'prSet', NS.dgm);
			return (
				point.getAttribute('type') === 'pres' &&
				props?.getAttribute('presAssocID') === picture.nodeId &&
				((picture.role && props.getAttribute('presName') === picture.role) ||
					(picture.styleLabel && props.getAttribute('presStyleLbl') === picture.styleLabel))
			);
		});
		if (presentations.length === 1) {
			const point = presentations[0]!;
			const props = first(point, 'spPr', NS.dgm) ?? doc.createElementNS(NS.dgm, 'dgm:spPr');
			if (!props.parentNode) point.appendChild(props);
			for (const old of [...props.childNodes]
				.filter((node): node is XmlElement => node.nodeType === 1)
				.filter(
					(node) =>
						node.namespaceURI === NS.a &&
						['blipFill', 'solidFill', 'gradFill', 'pattFill', 'noFill', 'grpFill'].includes(
							node.localName,
						),
				))
				props.removeChild(old);
			// Picture roles need their own direct fill. A content-node fill would paint captions too.
			props.insertBefore(doc.importNode(fill, true), props.firstChild);
		}
	}
	return buildXml(doc);
}

/** Relationship ids are local to the generated/data part, package targets are stable host identities. */
export function diagramPictureRelationships(
	owner: string,
	source: DiagramLayoutSource | undefined,
	existing?: string,
): string | undefined {
	if (!source?.pictures?.length || !source.pictureParts) return existing;
	const rels = parseRelationships(existing);
	for (const [id, part] of Object.entries(source.pictureParts)) {
		const previous = rels.get(id);
		if (
			previous &&
			(previous.mode === 'External' || resolvePartPath(owner, previous.target) !== part)
		)
			throw new Error(`Conflicting SmartArt picture relationship: ${id}`);
		rels.set(id, { id, type: RELATIONSHIP_TYPES.image, target: `/${part}`, mode: 'Internal' });
	}
	return buildRelationshipsXml(rels);
}
