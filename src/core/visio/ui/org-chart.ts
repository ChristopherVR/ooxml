import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioConnectorEdges } from './auto-layout';
import { visioShapeDataText } from './data-graphic-rules';

/**
 * The Shape Data rows of an org chart position, in the order and with the row names Visio's own
 * org chart masters (Executive, Manager, Position and Assistant Belt) carry, so a position keeps
 * its data when Visio opens the file. The Email row is labelled E-Mail, as Visio does.
 */
export const VISIO_ORG_ROWS = [
	{ row: 'Department', label: 'Department' },
	{ row: 'Telephone', label: 'Telephone' },
	{ row: 'Name', label: 'Name' },
	{ row: 'Title', label: 'Title' },
	{ row: 'Email', label: 'E-Mail' },
] as const;

export interface VisioOrgPositionData {
	Name?: string;
	Title?: string;
	Department?: string;
	Telephone?: string;
	Email?: string;
}

/**
 * The set-shape-data edits for the rows `data` names. The rows themselves come from the master the
 * position is an instance of; only the values given are written.
 */
export function visioOrgDataEdits(
	pageId: string,
	shapeId: string,
	data: VisioOrgPositionData,
): VisioEdit[] {
	return VISIO_ORG_ROWS.filter(({ row }) => data[row] !== undefined).map(({ row, label }) => ({
		type: 'set-shape-data',
		pageId,
		shapeId,
		row,
		data: { label, prompt: '', type: 'string', format: '', value: data[row] ?? '' },
	}));
}

/** Whether a shape is an org chart position: it carries every org chart Shape Data row. */
export const visioIsOrgPosition = (shape: VisioShape | undefined): boolean =>
	!!shape && VISIO_ORG_ROWS.every(({ row }) => shape.shapeData?.some((item) => item.name === row));

/** Whether a position is an instance of the Assistant Belt master (instances are named after it). */
export const visioIsOrgAssistant = (shape: VisioShape | undefined): boolean =>
	!!shape && /^Assistant Belt(\.\d+)?$/i.test(shape.name);

/** The org chart position a selected shape is, when it is one. */
export function visioOrgPosition(page: VisioPage, shapeId: string): VisioShape | undefined {
	const shape = page.shapes.find((item) => item.id === shapeId);
	return visioIsOrgPosition(shape) ? shape : undefined;
}

export interface VisioOrgHierarchy {
	/** Every position on the page. */
	positions: VisioShape[];
	/** Reporting lines between positions, manager to subordinate. */
	lines: { from: string; to: string }[];
	parentOf(id: string): string | undefined;
	childrenOf(id: string): string[];
}

/** The positions of a page and the glued connectors that join them. */
export function visioOrgHierarchy(page: VisioPage): VisioOrgHierarchy {
	const positions = page.shapes.filter(visioIsOrgPosition);
	const ids = new Set(positions.map((shape) => shape.id));
	const lines = visioConnectorEdges(page).filter((edge) => ids.has(edge.from) && ids.has(edge.to));
	return {
		positions,
		lines,
		parentOf: (id) => lines.find((line) => line.to === id)?.from,
		childrenOf: (id) => lines.filter((line) => line.from === id).map((line) => line.to),
	};
}

/** The text a position shows: its Name, then its Title, each on a line of its own when set. */
export const visioOrgText = (data: Pick<VisioOrgPositionData, 'Name' | 'Title'>): string =>
	[data.Name, data.Title].filter((line) => line?.trim()).join('\n');

/**
 * Keep the text of an org chart position in step with its Shape Data, as Visio's org chart shows
 * the Name and Title fields: after a `set-shape-data` of the Name or Title row of a position, a
 * text edit with the new Name and Title follows it. Other edits are returned as they are.
 */
export function visioOrgFollowData(page: VisioPage, edits: readonly VisioEdit[]): VisioEdit[] {
	const result: VisioEdit[] = [...edits];
	const seen = new Set<string>();
	for (const edit of edits) {
		if (edit.type !== 'set-shape-data' || (edit.row !== 'Name' && edit.row !== 'Title')) continue;
		if (edit.pageId !== page.id || seen.has(edit.shapeId)) continue;
		const shape = visioOrgPosition(page, edit.shapeId);
		if (!shape) continue;
		seen.add(edit.shapeId);
		// The last edit of a row in this batch wins, as the edits are applied in order.
		const value = (row: 'Name' | 'Title'): string => {
			const last = [...edits]
				.reverse()
				.find(
					(item) =>
						item.type === 'set-shape-data' &&
						item.shapeId === edit.shapeId &&
						item.pageId === page.id &&
						item.row === row,
				);
			if (last?.type === 'set-shape-data') return last.data?.value ?? '';
			return visioShapeDataText(shape, row) ?? '';
		};
		result.push({
			type: 'replace-plain-text',
			pageId: page.id,
			shapeId: edit.shapeId,
			text: visioOrgText({ Name: value('Name'), Title: value('Title') }),
		});
	}
	return result;
}
