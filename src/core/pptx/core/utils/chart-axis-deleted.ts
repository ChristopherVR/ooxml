/**
 * Reconcile a loaded axis's `c:delete` (`CT_CatAx` / `CT_ValAx` / `CT_DateAx` /
 * `CT_SerAx`: `axId`, `scaling`, `delete`, `axPos`, ...) with
 * `PptxChartAxisFormatting.deleted`.
 *
 * The parser reports `deleted: true` for `c:delete val="1"` and nothing
 * otherwise, so `undefined` means "no edit" here and the element is left as
 * authored; an explicit `true` / `false` is written only when it differs.
 *
 * @module utils/chart-axis-deleted
 */
import type { XmlObject } from '../types';

type GetLocalName = (key: string) => string;

function authoredDeleted(node: unknown): boolean {
	if (node === undefined) {
		return false;
	}
	const val = node && typeof node === 'object' ? (node as XmlObject)['@_val'] : undefined;
	if (val === undefined || val === null || val === '') {
		return true;
	}
	return !(val === '0' || val === 'false');
}

/** @returns Whether the axis changed. */
export function applyChartAxisDeletedToXml(
	axisNode: XmlObject,
	deleted: boolean | undefined,
	getLocalName: GetLocalName,
): boolean {
	if (deleted === undefined) {
		return false;
	}
	const key = Object.keys(axisNode).find((candidate) => getLocalName(candidate) === 'delete');
	if (authoredDeleted(key ? axisNode[key] : undefined) === deleted) {
		return false;
	}
	const node = { '@_val': deleted ? '1' : '0' };
	if (key) {
		axisNode[key] = node;
		return true;
	}
	const entries = Object.entries(axisNode);
	const at = entries.findIndex(([candidate]) => {
		const local = getLocalName(candidate);
		return local !== 'axId' && local !== 'scaling' && !candidate.startsWith('@_');
	});
	entries.splice(at === -1 ? entries.length : at, 0, ['c:delete', node]);
	for (const candidate of Object.keys(axisNode)) {
		delete axisNode[candidate];
	}
	for (const [candidate, value] of entries) {
		axisNode[candidate] = value;
	}
	return true;
}
