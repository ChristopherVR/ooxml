// Chart commands on the editor document: the chart a node selection is on, replacing it with an
// edited chart, and inserting a new one. Each returns a transaction the editor dispatches, so an
// edit is one step of the editor's undo history (the chart rides on the image node as JSON).
import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { DocxChart } from '../chart';
import { isDocxChartEditable } from '../chart-edit';
import type { InlineImage } from '../inline-model';
import { nodeChart } from './chart-spec';

/** A selected chart: its position, node and model. */
export interface SelectedChart {
	pos: number;
	node: ProseMirrorNode;
	chart: DocxChart;
	/** A classic chart part that loaded (extended charts stay read-only). */
	editable: boolean;
}

/** The chart the selection is on (a node selection of an image that carries a chart), if any. */
export function selectedChart(state: EditorState): SelectedChart | undefined {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || selection.node.type.name !== 'image')
		return undefined;
	const chart = nodeChart(selection.node);
	if (!chart) return undefined;
	return { pos: selection.from, node: selection.node, chart, editable: isDocxChartEditable(chart) };
}

/** Replaces the selected chart's model with `chart`, keeping the chart selected; one undo step. */
export function replaceSelectedChart(state: EditorState, chart: DocxChart): Transaction | null {
	const selected = selectedChart(state);
	if (!selected) return null;
	const tr = state.tr.setNodeMarkup(selected.pos, undefined, {
		...selected.node.attrs,
		chart: JSON.stringify(chart),
	});
	return closeHistory(tr.setSelection(NodeSelection.create(tr.doc, selected.pos)));
}

/** Inserts the chart drawing `image` at the selection and selects it; one undo step. */
export function insertChartImage(state: EditorState, image: InlineImage): Transaction | null {
	const type = state.schema.nodes.image;
	if (!type || !image.chart) return null;
	const node = type.create({
		relId: image.relId,
		partName: image.partName,
		contentType: image.contentType,
		widthPx: image.widthPx,
		heightPx: image.heightPx,
		altText: image.altText ?? null,
		title: image.title ?? null,
		anchored: false,
		unsupported: image.unsupported ?? 'Chart',
		chart: JSON.stringify(image.chart),
	});
	const tr = closeHistory(state.tr.replaceSelectionWith(node, false));
	const pos = tr.mapping.map(state.selection.from, -1);
	const at = tr.doc.nodeAt(pos)?.type === type ? pos : tr.selection.from - node.nodeSize;
	if (at >= 0 && tr.doc.nodeAt(at)?.type === type)
		tr.setSelection(NodeSelection.create(tr.doc, at));
	return closeHistory(tr.scrollIntoView());
}
