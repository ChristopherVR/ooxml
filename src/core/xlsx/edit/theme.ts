import type { ThemePalette } from '../model';
import { DEFAULT_THEME } from '../workbook';
import type { EditContext } from './context';

/** A change to the workbook theme palette and/or scheme fonts. */
export interface ThemePatch {
	/**
	 * 12-colour palette in SpreadsheetML theme index order
	 * (`lt1`, `dk1`, `lt2`, `dk2`, `accent1`–`accent6`, `hlink`, `folHlink`).
	 * Entries may be 6-digit hex with or without a leading `#`.
	 */
	colors?: readonly string[];
	/** Major (heading) Latin theme font typeface. */
	majorFont?: string;
	/** Minor (body) Latin theme font typeface. */
	minorFont?: string;
}

const HEX_6 = /^#?([0-9A-Fa-f]{6})$/;

function normalizeHex(value: string | undefined, fallback: string): string {
	const match = value?.trim().match(HEX_6);
	return match?.[1] ? match[1].toUpperCase() : fallback;
}

function samePalette(a: ThemePalette, b: ThemePalette): boolean {
	return (
		a.majorFont === b.majorFont &&
		a.minorFont === b.minorFont &&
		a.colors.length === b.colors.length &&
		a.colors.every((color, index) => color === b.colors[index])
	);
}

/**
 * Updates the workbook theme's colour palette and/or scheme fonts as one undo step.
 * Replaces `workbook.theme` (never mutates it in place) so history snapshots restore cleanly.
 * Returns whether anything changed.
 */
export function setTheme(ctx: EditContext, patch: ThemePatch, label = 'Change theme'): boolean {
	const prior = ctx.workbook.theme;
	const colors = patch.colors
		? Array.from({ length: 12 }, (_, index) =>
				normalizeHex(
					patch.colors?.[index],
					prior.colors[index] ?? DEFAULT_THEME.colors[index] ?? '000000',
				),
			)
		: [...prior.colors];
	const majorFont = patch.majorFont?.trim() || prior.majorFont;
	const minorFont = patch.minorFont?.trim() || prior.minorFont;
	const next: ThemePalette = { ...prior, colors, majorFont, minorFont };
	if (samePalette(next, prior)) return false;
	ctx.run(label, 'view', [{ kind: 'meta' }], () => {
		ctx.workbook.theme = next;
	});
	return true;
}
