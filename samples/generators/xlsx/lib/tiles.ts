/** Headline tiles of a dashboard: a small label, a large figure and a line of context. */
import type { Look } from './book';
import type { Sheet } from './sheet';

export interface Tile {
	/** The first and last column the tile spans. */
	first: string;
	last: string;
	/** The label: plain text, or a formula when it starts with `=`. */
	label: string;
	/** The formula of the figure, and its number format. */
	value: string;
	fmt: string;
	/** The formula of the line under the figure. */
	note: string;
	/** The accent of this tile, when it differs from the others. */
	color?: string;
}

export interface TileColors {
	fill: string;
	accent: string;
	muted: string;
}

/** Writes the tiles on three rows starting at `top`. */
export function tiles(sheet: Sheet, top: number, specs: Tile[], colors: TileColors): void {
	// A white edge on the right separates neighbouring tiles without spending a column.
	const gap: Look = { right: ['thick', 'FFFFFF'] };
	const [label, figure, note] = [top, top + 1, top + 2];
	for (const spec of specs) {
		const accent = spec.color ?? colors.accent;
		const body: Look = { fill: colors.fill, indent: 1 };
		const head: Look = { ...body, top: ['medium', accent] };
		sheet
			.paint(`${spec.first}${label}:${spec.last}${label}`, head)
			.paint(`${spec.first}${figure}:${spec.last}${note}`, body)
			.paint(`${spec.last}${label}`, { ...head, ...gap })
			.paint(`${spec.last}${figure}:${spec.last}${note}`, { ...body, ...gap });
		const small: Look = { color: colors.muted, size: 9 };
		if (spec.label.startsWith('='))
			sheet.formula(`${spec.first}${label}`, spec.label.slice(1), {
				...head,
				...small,
				v: 'bottom',
			});
		else sheet.set(`${spec.first}${label}`, spec.label, { ...head, ...small, v: 'bottom' });
		sheet.formula(`${spec.first}${figure}`, spec.value, {
			...body,
			font: 'display',
			size: 20,
			bold: true,
			color: accent,
			fmt: spec.fmt,
			h: 'left',
			v: 'center',
		});
		sheet.formula(`${spec.first}${note}`, spec.note, { ...body, ...small, v: 'top' });
		// The figure needs the whole tile: a number does not overflow into the next cell.
		sheet.merge(`${spec.first}${figure}:${spec.last}${figure}`);
	}
	sheet.heights({ [label]: 20, [figure]: 32, [note]: 20 });
}
