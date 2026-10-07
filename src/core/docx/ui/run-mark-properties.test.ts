import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import type { TextRun } from '../model';
import { marksForRun } from './run-marks';
import { applyMarkFormatting, linkFromMarks } from './run-mark-properties';
import { markSpecs } from './schema-marks';

const schema = new Schema({ nodes: { doc: { content: 'text*' }, text: {} }, marks: markSpecs });

describe('shared mark-to-run conversion', () => {
	it('round-trips overlapping revision history, script fonts, explicit off, links and comments with an independent schema', () => {
		const original: TextRun = {
			text: 'Text',
			bold: false,
			italic: true,
			fontFamilyComplexScript: 'Amiri',
			fontFamilyEastAsia: 'Yu Mincho',
			fontSizeComplexScript: 14.5,
			language: 'ar',
			bidiLanguage: 'ar',
			rtl: true,
			style: 'Character',
			revision: {
				kind: 'insert',
				id: 'text',
				author: 'Ada',
				date: '2026-10-07T00:00:00Z',
				dateUtc: '2026-10-07T00:00:00Z',
			},
			formatRevision: {
				kind: 'formatChange',
				id: 'format',
				author: 'Grace',
				previousRunPropertiesXml:
					'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
			},
			link: { href: 'https://example.com/', tooltip: 'Example' },
			commentIds: ['a', 'b'],
		};
		const node = schema.text(original.text, marksForRun(original, schema));
		const run: TextRun = { text: node.text! };
		const link = linkFromMarks(node);
		if (link) run.link = link;
		applyMarkFormatting(run, node);
		expect(run).toEqual(original);
		// Model snapshots are independent of the editor mark's object identity.
		run.formatRevision!.author = 'Changed';
		expect(original.formatRevision?.author).toBe('Grace');
	});

	it('retains applied bold while an opaque mark carries an inherited explicit off', () => {
		const original: TextRun = { text: 'Text', bold: false, fontFamilyComplexScript: 'Amiri' };
		const node = schema.text('Text', [
			...marksForRun(original, schema),
			schema.marks.bold!.create(),
		]);
		const run: TextRun = { text: 'Text' };
		applyMarkFormatting(run, node);
		expect(run).toEqual({ text: 'Text', bold: true, fontFamilyComplexScript: 'Amiri' });
	});
});
