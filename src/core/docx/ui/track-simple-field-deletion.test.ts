import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { markSpecs } from './schema-marks';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';
import { fieldGuardPlugin } from './field-guard';
import { deleteSimpleFieldResult } from './simple-field-input';
import { trackChangesPlugin } from './track-changes-mode';
import { acceptAllChanges, rejectAllChanges, collectRevisionRanges } from './review-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
for (const ownInsertion of [false, true])
	for (const action of ['accept', 'reject'])
		it(`${action} resolves whole simple-result deletion with one field (${ownInsertion ? 'own insertion' : 'imported text'})`, () => {
			const result = runToInlineNodes(
				{
					text: 'AB',
					bold: true,
					field: { instr: 'QUOTE "AB"', simple: true },
					fieldInstanceId: 'first',
					...(ownInsertion
						? { revision: { kind: 'insert' as const, author: 'Ada', id: 'prior' } }
						: {}),
				},
				schema,
			);
			const doc = schema.node(
				'doc',
				null,
				schema.node('paragraph', null, [schema.text('L'), ...result, schema.text('R')]),
			);
			let state = EditorState.create({
				doc,
				selection: TextSelection.create(doc, 2, 4),
				plugins: [
					fieldGuardPlugin(),
					trackChangesPlugin(
						() => 'Ada',
						() => true,
					),
				],
			});
			state = state.applyTransaction(deleteSimpleFieldResult(state, true)!).state;
			const runs = () =>
				Array.from({ length: state.doc.firstChild!.childCount }, (_, index) =>
					inlineNodeRun(state.doc.firstChild!.child(index))!,
				);
			expect(runs().filter((run) => run.fieldChar === 'begin')).toHaveLength(1);
			expect(runs().filter((run) => run.field?.simple)).toEqual([]);
			expect(collectRevisionRanges(state.doc)).toHaveLength(ownInsertion ? 0 : 1);
			if (!ownInsertion)
				expect(runs().find((run) => run.revision)).toMatchObject({
					text: 'AB',
					bold: true,
					field: { instr: 'QUOTE "AB"' },
					revision: { kind: 'delete', author: 'Ada' },
				});
			const view = {
				get state() {
					return state;
				},
				editable: true,
				dispatch(tr: import('prosemirror-state').Transaction) {
					state = state.applyTransaction(tr).state;
				},
			} as EditorView;
			if (!ownInsertion) {
				(action === 'accept' ? acceptAllChanges : rejectAllChanges)(view);
				expect(state.doc.textContent).toBe(action === 'accept' ? 'LR' : 'LABR');
			} else expect(state.doc.textContent).toBe('LR');
			expect(runs().filter((run) => run.fieldChar === 'begin')).toHaveLength(1);
		});
