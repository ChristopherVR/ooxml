import { history, undo } from 'prosemirror-history';
import { Schema } from 'prosemirror-model';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { describe, expect, it } from 'vitest';
import { setChartType } from '../../chart/index';
import { editDocxChart, newChartImage, newDocxChart } from '../chart-edit';
import { insertChartImage, replaceSelectedChart, selectedChart } from './chart-commands';
import { imageNodeSpec } from './inline-content-schema';
import { markSpecs } from './schema-marks';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0] },
		text: { group: 'inline' },
		image: imageNodeSpec,
	},
	marks: markSpecs,
});

function emptyState(): EditorState {
	const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('Hello')])]);
	return EditorState.create({
		doc,
		plugins: [history()],
		selection: TextSelection.create(doc, 6),
	});
}

describe('chart commands', () => {
	it('inserts a chart, selects it and edits it as one undo step each', () => {
		let state = emptyState();
		state = state.apply(insertChartImage(state, newChartImage(newDocxChart()))!);
		const inserted = selectedChart(state);
		expect(state.selection).toBeInstanceOf(NodeSelection);
		expect(inserted?.editable).toBe(true);
		expect(inserted?.node.attrs.widthPx).toBe(576);
		const pie = editDocxChart(inserted!.chart, (space) => setChartType(space, { kind: 'pie' }));
		state = state.apply(replaceSelectedChart(state, pie)!);
		expect(selectedChart(state)?.chart.chartSpace?.plotArea.groups[0]?.kind).toBe('pie');
		undo(state, (tr) => {
			state = state.apply(tr);
		});
		state = state.apply(state.tr.setSelection(NodeSelection.create(state.doc, 6)));
		expect(selectedChart(state)?.chart.chartSpace?.plotArea.groups[0]?.kind).toBe('bar');
		undo(state, (tr) => {
			state = state.apply(tr);
		});
		expect(state.doc.textContent).toBe('Hello');
		expect(selectedChart(state)).toBeUndefined();
	});

	it('does nothing without a selected chart', () => {
		expect(replaceSelectedChart(emptyState(), newDocxChart())).toBeNull();
	});
});
