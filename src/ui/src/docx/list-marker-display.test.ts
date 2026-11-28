// @vitest-environment jsdom
import { expect, it } from 'vitest';
import type { DocumentModel, Paragraph, ParagraphListLabel } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { listMarkerDisplay } from './list-marker-display';
import { schema } from './schema';

const model = {} as DocumentModel;
const paragraph = { runs: [{ text: 'Item' }] } as Paragraph;
const node = schema.nodes.paragraph!.create({}, schema.text('Item'));
const label = (extra: Partial<ParagraphListLabel>): ParagraphListLabel =>
	({
		numId: '1',
		level: 0,
		text: '•',
		suffix: 'tab',
		...extra,
	}) as ParagraphListLabel;
const textIndent = (css: string) => Number(/text-indent:(-?[\d.]+)px/.exec(css)?.[1]);

// Issue #371: the text after a legacy marker starts at its legacy indent, as in Print Layout.
it('starts the text at the legacy indent instead of the next tab stop', () => {
	const display = listMarkerDisplay(
		node,
		paragraph,
		label({ legacy: { enabled: true, indentTwips: 283, spaceTwips: 0 } } as never),
		model,
		createFakeMeasurer(),
	);
	expect(textIndent(display.css)).toBeCloseTo(283 / 15, 1);
});

it('measures the legacy indent from the number start and keeps legacySpace after it', () => {
	const display = listMarkerDisplay(
		node,
		{ ...paragraph, indentLeftTwips: 1440 } as Paragraph,
		label({ legacy: { enabled: true, indentTwips: 283, spaceTwips: 0 } } as never),
		model,
		createFakeMeasurer(),
	);
	// The number starts at the 96px indent, so the text is 18.87px past it (text-indent is
	// relative to the indent).
	expect(textIndent(display.css)).toBeCloseTo(283 / 15, 1);
	const spaced = listMarkerDisplay(
		node,
		paragraph,
		label({ legacy: { enabled: true, indentTwips: 0, spaceTwips: 300 } } as never),
		model,
		createFakeMeasurer(),
	);
	expect(textIndent(spaced.css)).toBeGreaterThanOrEqual(300 / 15);
});
