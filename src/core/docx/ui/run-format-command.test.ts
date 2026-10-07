import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { loadDocx } from '../parse';
import type { Paragraph, TextRun } from '../model';
import { signedTwips, halfPoints } from '../units';
import { markSpecs } from './schema-marks';
import { imageNodeSpec } from './inline-content-schema';
import {
	hardBreakNodeSpec,
	pageBreakNodeSpec,
	noteReferenceNodeSpec,
	fieldMarkerNodeSpec,
} from './break-note-schema';
import { runToInlineNodes, appendInlineNode } from './run-adapter';
import { applyRunFormattingPatch } from './run-format-command';
import { trackChangesPlugin, trackChangesPluginKey } from './track-changes-mode';
import { resolveFormattingRange } from './review-formatting';
import { collectRevisionRanges } from './review-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		image: imageNodeSpec,
		hardBreak: hardBreakNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
const content = (paragraphs: Paragraph[]) =>
	paragraphs.map((p) =>
		p.runs.map(
			({ sourceRunPropertiesXml: _source, restoredRunPropertiesXml: _restored, ...run }) => run,
		),
	);

for (const name of ['picture', 'note', 'break', 'field', 'line-break'])
	it(`records and resolves native ${name} advanced properties without changing inline content`, async () => {
		const fixture = (mode: string) =>
			readFile(
				new URL(
					`../__fixtures__/review-advanced-object-formatting/${name}-${mode}.docx`,
					import.meta.url,
				),
			);
		const loaded = await loadDocx(await fixture('before'));
		const paragraphs = loaded.model.blocks as Paragraph[];
		let editor = EditorState.create({
			doc: schema.node(
				'doc',
				null,
				paragraphs.map((p) =>
					schema.node(
						'paragraph',
						null,
						p.runs.flatMap((run) => runToInlineNodes(run, schema)),
					),
				),
			),
			plugins: [
				history(),
				trackChangesPlugin(
					() => 'Ada',
					() => true,
				),
			],
		});
		let position = -1;
		editor.doc.descendants((node, pos) => {
			if (
				position >= 0 ||
				!node.isInline ||
				node.isText ||
				(node.type.name === 'fieldMarker' && node.attrs.kind !== 'code')
			)
				return;
			position = pos;
		});
		expect(position).toBeGreaterThan(0);
		const initial = editor.doc;
		editor = editor.apply(
			editor.tr.setSelection(TextSelection.create(editor.doc, position, position + 1)),
		);
		applyRunFormattingPatch({
			fontSize: 18,
			color: '#C00000',
			smallCaps: true,
			characterSpacingTwips: signedTwips(30),
			textScalePercent: 150,
			positionHalfPoints: halfPoints(4),
			kerningHalfPoints: halfPoints(24),
		})(editor, (tr) => {
			editor = editor.apply(tr);
		});
		expect(collectRevisionRanges(editor.doc)).toMatchObject([
			{ kind: 'formatChange', author: 'Ada' },
		]);
		for (const mode of ['accept', 'reject'] as const) {
			const tr = editor.tr;
			resolveFormattingRange(tr, position, position + 1, mode);
			const resolved = editor.apply(tr.setMeta(trackChangesPluginKey, { tracked: true }));
			const model = {
				...loaded.model,
				blocks: paragraphs.map((p, index) => {
					const runs: TextRun[] = [];
					resolved.doc.child(index).forEach((node) => appendInlineNode(runs, node));
					return { ...p, runs };
				}),
			};
			const exported = (await loadDocx(await loaded.save(model))).model;
			const native = (await loadDocx(await fixture(mode === 'accept' ? 'accepted' : 'rejected')))
				.model;
			expect(content(exported.blocks as Paragraph[])).toEqual(
				content(native.blocks as Paragraph[]),
			);
		}
		expect(
			undo(editor, (tr) => {
				editor = editor.apply(tr);
			}),
		).toBe(true);
		expect(editor.doc.eq(initial)).toBe(true);
	});
