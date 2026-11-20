// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// How a simple text box (the form `write-text-box.ts` writes and patches) is drawn after an edit:
// its frame and paragraphs, so the editor shows what saving will write.
import type { DrawingFill } from '../drawingml/index';
import type { InlineImage, Paragraph } from './model';
import type { ShapeSpec } from './shape-model';

const solid = (value: string): DrawingFill => ({
	kind: 'solid',
	color: { kind: 'srgb', value, transforms: [] },
});

/** The outline `write-text-box.ts` writes: 0.75 pt black, or none. */
const outline = (border: boolean): NonNullable<ShapeSpec['line']> =>
	border ? { widthEmu: 9525, fill: solid('000000') } : { fill: { kind: 'none' } };

/**
 * The frame and paragraphs of a simple text box holding `lines` with or without an outline. What
 * the box had before (`previous`) is kept where the edit does not change it: its frame when the
 * outline is unchanged, its paragraphs, insets and anchor when the lines are.
 */
export function simpleTextBoxDrawing(
	lines: readonly string[],
	border: boolean,
	previous: Pick<InlineImage, 'shape' | 'textBoxContent' | 'textBoxText' | 'textBoxBorder'> = {},
): Required<Pick<InlineImage, 'shape' | 'textBoxContent'>> {
	const frame: ShapeSpec = previous.shape ?? {
		preset: 'rect',
		fill: solid('FFFFFF'),
		line: outline(border),
		editable: false,
		textEditable: false,
	};
	const shape =
		previous.shape && (previous.textBoxBorder !== false) === border
			? frame
			: { ...frame, line: outline(border) };
	const sameLines =
		previous.textBoxContent && JSON.stringify(previous.textBoxText ?? []) === JSON.stringify(lines);
	const blocks: Paragraph[] = (lines.length ? lines : ['']).map((line, index) => ({
		type: 'paragraph',
		id: `txbx-p${index}`,
		runs: [{ text: line }],
	}));
	return {
		shape,
		textBoxContent: sameLines
			? previous.textBoxContent!
			: {
					anchor: 't',
					insetLeftEmu: 91440,
					insetTopEmu: 45720,
					insetRightEmu: 91440,
					insetBottomEmu: 45720,
					...previous.textBoxContent,
					blocks,
				},
	};
}
