import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import type { TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { imageNodeSpec } from './inline-content-schema';
import { fieldMarkerNodeSpec, noteReferenceNodeSpec, pageBreakNodeSpec } from './break-note-schema';
import { appendInlineNode, inlineNodeRun, runToInlineNodes } from './run-adapter';

const schema = new Schema({
	nodes: {
		doc: { content: 'inline*' },
		text: { group: 'inline' },
		image: imageNodeSpec,
		pageBreak: pageBreakNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
		noteReference: noteReferenceNodeSpec,
		hardBreak: { inline: true, group: 'inline' },
		equation: {
			inline: true,
			group: 'inline',
			attrs: { omml: {}, display: {}, format: { default: null } },
		},
	},
	marks: markSpecs,
});
const atoms: TextRun[] = [
	{ text: '', break: 'page' },
	{ text: '', break: 'column' },
	{ text: '', noteReference: { kind: 'footnote', id: '1' } },
	{ text: '', fieldChar: 'begin' },
	{ text: '', fieldCode: ' PAGE ' },
	{ text: '', equation: { omml: '<m:oMath/>', display: false } },
	{
		text: '',
		image: {
			relId: 'rId1',
			partName: 'word/media/picture.png',
			contentType: 'image/png',
			widthPx: 20,
			heightPx: 30,
		},
		link: { href: 'https://example.com' },
	},
];
for (const atom of atoms)
	it(`preserves ${atom.break ?? (atom.image ? 'picture' : atom.equation ? 'equation' : atom.noteReference ? 'note reference' : (atom.fieldChar ?? 'field code'))} properties and overlapping histories with an independent schema`, () => {
		const run: TextRun = {
			...atom,
			bold: false,
			language: 'en-GB',
			fontFamilyComplexScript: 'Amiri',
			sourceRunPropertiesXml: '<w:rPr><w:custom/></w:rPr>',
			commentIds: ['comment'],
			revision: { kind: 'moveTo', id: 'move', author: 'Ada', move: { name: 'move' } },
			formatRevision: {
				kind: 'formatChange',
				id: 'format',
				author: 'Bob',
				previousRunPropertiesXml: '<w:rPr/>',
			},
		};
		const original = structuredClone(run);
		const nodes = runToInlineNodes(run, schema);
		expect(nodes).toHaveLength(1);
		expect(nodes[0]!.nodeSize).toBe(1);
		expect(inlineNodeRun(nodes[0]!)).toEqual(run);
		expect(run).toEqual(original);
	});

it('keeps histories and property bases through text and hard-break splits', () => {
	const run: TextRun = {
		text: 'A\nB',
		bold: true,
		sourceRunPropertiesXml: '<w:rPr><w:b/></w:rPr>',
		revision: { id: 'r', kind: 'delete', author: 'Ada' },
	};
	const runs: TextRun[] = [];
	for (const node of runToInlineNodes(run, schema)) appendInlineNode(runs, node);
	expect(runs).toEqual([run]);
});
