import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { hardBreakNodeSpec } from './break-note-schema';
import { markSpecs } from './schema-marks';
import { runToInlineNodes, inlineNodeRun } from './run-adapter';
import { trackChangesPlugin } from './track-changes-mode';
import { collectRevisionRanges } from './review-commands';
import { clearInlineTextRevisions } from './review-inline-revisions';

for (const modern of [true, false]) {
	const schema = new Schema({
		nodes: {
			doc: { content: 'paragraph+' },
			paragraph: { content: 'inline*' },
			text: { group: 'inline' },
			hardBreak: modern
				? hardBreakNodeSpec
				: { group: 'inline', inline: true, leafText: () => '\n' },
		},
		marks: markSpecs,
	});
	it(`links line-break cut/paste revisions and retains formatting (${modern ? 'attribute' : 'legacy'} schema)`, () => {
		let editor = EditorState.create({
			doc: schema.node(
				'doc',
				null,
				schema.node(
					'paragraph',
					null,
					runToInlineNodes({ text: 'A\nB', italic: true, language: 'en-GB' }, schema),
				),
			),
			plugins: [
				trackChangesPlugin(
					() => 'Ada',
					() => true,
				),
			],
		});
		editor = editor.apply(editor.tr.delete(2, 3).setMeta('uiEvent', 'cut'));
		expect(collectRevisionRanges(editor.doc)).toMatchObject([{ kind: 'delete' }]);
		editor = editor.apply(
			editor.tr.insert(4, schema.node('hardBreak')).setMeta('uiEvent', 'paste'),
		);
		const ranges = collectRevisionRanges(editor.doc);
		expect(ranges).toHaveLength(2);
		expect(ranges[0]!.move).toBeTruthy();
		expect(ranges[1]!.move).toBe(ranges[0]!.move);
		const retained = inlineNodeRun(editor.doc.nodeAt(2)!)!;
		expect(retained).toMatchObject({
			text: '\n',
			italic: true,
			language: 'en-GB',
			revision: { kind: 'moveFrom' },
		});
		if (modern) expect(editor.doc.nodeAt(2)!.marks).toHaveLength(0);
		const tr = editor.tr;
		clearInlineTextRevisions(tr, 2, 3);
		tr.removeMark(2, 3, schema.marks.deletion);
		expect(inlineNodeRun(tr.doc.nodeAt(2)!)).toMatchObject({
			text: '\n',
			italic: true,
			language: 'en-GB',
		});
		expect(inlineNodeRun(tr.doc.nodeAt(2)!)!.revision).toBeUndefined();
	});
}
