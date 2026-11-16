// Binding a regenerated SmartArt cache to the presentation points a source cache was laid out with.
import { NS, children, first, type XmlElement } from '../../xml/index';

const PRESENTATION_GUID = /^\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$/u;

/**
 * The presentation points Office laid out (or another producer wrote), keyed by the content node
 * they present and their layout role: a regenerated shape for the same node and role takes the
 * point's id, as Office's own cache binds `dsp:sp/@modelId` to it, instead of a parallel point.
 * Points this writer generated before are excluded; they are replaced.
 */
export function sourcePresentationIds(
	points: XmlElement,
	generated: ReadonlySet<string | null>,
	shapes: readonly { modelId: string; sourceNodeId?: string; sourceLayoutRole?: string }[],
): Map<string, string> {
	const byRole = new Map<string, string[]>();
	for (const point of children(points, 'pt', NS.dgm)) {
		const id = point.getAttribute('modelId');
		if (point.getAttribute('type') !== 'pres' || !id || generated.has(id)) continue;
		const set = first(point, 'prSet', NS.dgm);
		const node = set?.getAttribute('presAssocID');
		const role = set?.getAttribute('presName');
		if (!node || !role || !PRESENTATION_GUID.test(id)) continue;
		const key = `${node}|${role}`;
		byRole.set(key, [...(byRole.get(key) ?? []), id]);
	}
	const ids = new Map<string, string>();
	for (const shape of shapes) {
		if (!shape.sourceNodeId || !shape.sourceLayoutRole) continue;
		const id = byRole.get(`${shape.sourceNodeId}|${shape.sourceLayoutRole}`)?.shift();
		if (id) ids.set(shape.modelId, id);
	}
	return ids;
}
