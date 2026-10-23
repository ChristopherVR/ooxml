/**
 * Keyword resolution of a raw layout type string to the layout family a
 * reflow uses. Moved from `pptx/core/utils/smartart-editing-reflow.ts`.
 *
 * @module diagram/edit/reflow-category
 */

import type { DiagramLayoutType } from '../types';

/**
 * Resolve a raw layout type string to a `DiagramLayoutType`.
 *
 * Matches keywords in the raw string to determine the most appropriate
 * layout category. Falls back to `"list"` when no match is found.
 *
 * @param layoutType - Raw layout type string from the PPTX data model.
 * @returns Resolved SmartArt layout type.
 */
export function resolveDiagramReflowCategory(layoutType: string | undefined): DiagramLayoutType {
	if (!layoutType) {
		return 'list';
	}
	const lower = layoutType.toLowerCase();

	if (lower.includes('hierarchy') || lower.includes('org')) {
		return 'hierarchy';
	}
	if (lower.includes('cycle') || lower.includes('radial')) {
		return 'cycle';
	}
	if (lower.includes('funnel')) {
		return 'funnel';
	}
	if (lower.includes('target') || lower.includes('bullseye')) {
		return 'target';
	}
	if (lower.includes('gear') || lower.includes('interlock')) {
		return 'gear';
	}
	if (lower.includes('venn')) {
		return 'venn';
	}
	if (lower.includes('timeline')) {
		return 'timeline';
	}
	if (lower.includes('bending') || lower.includes('snake') || lower.includes('zigzag')) {
		return 'bending';
	}
	if (lower.includes('chevron')) {
		return 'chevron';
	}
	if (lower.includes('relationship')) {
		return 'relationship';
	}
	if (lower.includes('process') || lower.includes('arrow')) {
		return 'process';
	}
	if (lower.includes('matrix')) {
		return 'matrix';
	}
	if (lower.includes('pyramid')) {
		return 'pyramid';
	}
	if (lower.includes('list') || lower.includes('block')) {
		return 'list';
	}

	return 'list';
}
