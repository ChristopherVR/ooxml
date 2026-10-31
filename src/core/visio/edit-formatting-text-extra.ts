import { attribute } from './sheet';
import { fail } from './package-common';
import { visioFormulaCachedValue } from './formula';
import { formattingCell } from './edit-style-admission';
import {
	VISIO_TEXT_CASES,
	VISIO_TEXT_POSITIONS,
	textExtraCharacterKeys,
	textExtraParagraphKeys,
	type VisioTextFormatExtras,
} from './edit-formatting-text-commands';

export type AddTextWrite = (
	name: string,
	value: string | number,
	unit?: string,
	formula?: string,
	dropUnit?: boolean,
) => void;

export const hasCharacterExtras = (edit: VisioTextFormatExtras): boolean =>
	textExtraCharacterKeys.some((key) => edit[key] !== undefined);
export const hasParagraphExtras = (edit: VisioTextFormatExtras): boolean =>
	textExtraParagraphKeys.some((key) => edit[key] !== undefined);

const quoted = (text: string) => `"${text.replaceAll('"', '""')}"`;
const rgb = (color: string) =>
	`RGB(${[1, 3, 5].map((i) => parseInt(color.slice(i, i + 2), 16)).join(',')})`;
/** Round cached inches so repeated edits compare equal and stay readable. */
const inches = (value: number) => Number(value.toFixed(10));

/** Native cell values: https://learn.microsoft.com/en-us/office/client-developer/visio/character-section */
export function characterExtraWrites(
	edit: VisioTextFormatExtras,
	name: (cell: string) => string,
	add: AddTextWrite,
): void {
	if (edit.fontTransparency !== undefined) add(name('ColorTrans'), edit.fontTransparency / 100);
	if (edit.textCase !== undefined) add(name('Case'), VISIO_TEXT_CASES.indexOf(edit.textCase));
	if (edit.textPosition !== undefined)
		add(name('Pos'), VISIO_TEXT_POSITIONS.indexOf(edit.textPosition));
	if (edit.language !== undefined) add(name('LangID'), edit.language);
	if (edit.letterSpacing !== undefined)
		add(name('Letterspace'), inches(edit.letterSpacing / 72), 'PT');
}

/** https://learn.microsoft.com/en-us/office/client-developer/visio/paragraph-section */
export function paragraphExtraWrites(
	edit: VisioTextFormatExtras,
	name: (cell: string) => string,
	add: AddTextWrite,
): void {
	for (const [key, cell] of [
		['indentRight', 'IndRight'],
		['indentFirst', 'IndFirst'],
		['spaceBefore', 'SpBefore'],
		['spaceAfter', 'SpAfter'],
	] as const)
		if (edit[key] !== undefined) add(name(cell), inches(edit[key] / 72), 'PT');
	if (edit.lineSpacing !== undefined) {
		// Negative SpLine values are multiples of the font size; positive values are exact.
		if (edit.lineSpacing.kind === 'multiple')
			add(name('SpLine'), -edit.lineSpacing.value, undefined, undefined, true);
		else add(name('SpLine'), inches(edit.lineSpacing.value / 72), 'PT');
	}
	if (edit.bulletStyle !== undefined) add(name('Bullet'), edit.bulletStyle);
	if (edit.bulletText !== undefined)
		add(name('BulletStr'), edit.bulletText, undefined, quoted(edit.bulletText));
}

const TAB_ALIGNMENTS = ['left', 'center', 'right', 'decimal'] as const;

/**
 * Tabs row 0 as Visio 16 saves it: `Position<n>` (inches) and `Alignment<n>` pairs counted from 0.
 * Visio 16 ignored the alignment of a row that held a single pair, so at least two pairs are
 * always written, padded (like the pair Visio adds itself) with a stop at 0 that never matches.
 * Pairs the shape already had beyond the new list are zeroed, since a cell cannot be removed here.
 */
export function tabStopWrites(
	stops: readonly { position: number; align: (typeof TAB_ALIGNMENTS)[number] }[],
	context: { source: ReadonlyMap<string, ReadonlyMap<string, unknown>> },
	add: AddTextWrite,
): void {
	const existing = [...(context.source.get('0')?.keys() ?? [])]
		.map((key) => /^Position(\d+)$/.exec(key))
		.reduce((count, match) => (match ? Math.max(count, Number(match[1]) + 1) : count), 0);
	const count = Math.max(2, stops.length, existing);
	for (let index = 0; index < count; index++) {
		const stop = stops[index];
		add(`Tabs.0.Position${index}`, stop ? inches(stop.position / 72) : 0, 'IN');
		add(`Tabs.0.Alignment${index}`, stop ? TAB_ALIGNMENTS.indexOf(stop.align) : 0);
	}
}

function shapeDimension(shape: Element, cell: 'Width' | 'Height'): number {
	const source = formattingCell(shape, cell);
	let value: number | undefined;
	try {
		const cached = visioFormulaCachedValue(attribute(source, 'V') ?? '', attribute(source, 'U'));
		if (cached.unit === 'length' || cached.unit === 'scalar') value = cached.value;
	} catch {
		/* Refused below. */
	}
	if (value === undefined || !(value > 0))
		fail('UNSUPPORTED_FORMAT_EDIT', `The text block needs a cached local ${cell}.`);
	return value;
}

/** Text Block Format and Text Transform cells on the shape sheet itself. */
export function textBlockWrites(
	shape: Element,
	edit: VisioTextFormatExtras,
	add: AddTextWrite,
): void {
	if (edit.margins)
		for (const [side, cell] of [
			['left', 'LeftMargin'],
			['right', 'RightMargin'],
			['top', 'TopMargin'],
			['bottom', 'BottomMargin'],
		] as const) {
			const value = edit.margins[side];
			if (value !== undefined) add(cell, inches(value / 72), 'PT');
		}
	if (edit.textBackground !== undefined) {
		if (edit.textBackground === 'none') add('TextBkgnd', 0);
		// Saved RGB backgrounds carry Visio's +1 palette offset in the formula only.
		else add('TextBkgnd', edit.textBackground, undefined, `${rgb(edit.textBackground)}+1`);
	}
	if (edit.textBackgroundTransparency !== undefined)
		add('TextBkgndTrans', edit.textBackgroundTransparency / 100);
	const block = edit.textBlock;
	if (block) {
		const width = shapeDimension(shape, 'Width'),
			height = shapeDimension(shape, 'Height');
		const [x, y, w, h] = [block.x, block.y, block.width, block.height].map((value) =>
			Number(value.toFixed(6)),
		) as [number, number, number, number];
		const textWidth = inches(width * w),
			textHeight = inches(height * h);
		// Proportional formulas keep the text block in place when the shape is resized.
		add('TxtPinX', inches(width * x), 'IN', `Width*${x}`);
		add('TxtPinY', inches(height * y), 'IN', `Height*${y}`);
		add('TxtWidth', textWidth, 'IN', `Width*${w}`);
		add('TxtHeight', textHeight, 'IN', `Height*${h}`);
		add('TxtLocPinX', inches(textWidth / 2), 'IN', 'TxtWidth*0.5');
		add('TxtLocPinY', inches(textHeight / 2), 'IN', 'TxtHeight*0.5');
		add('TxtAngle', inches(block.angle), 'DEG');
	}
}
