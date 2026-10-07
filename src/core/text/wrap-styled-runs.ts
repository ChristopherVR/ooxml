export interface WrapStyledRunsOptions {
	/** Break an oversized word at grapheme boundaries. Otherwise it overflows intact. */
	breakWords?: boolean;
}

interface Piece<T> {
	run: T;
	source: number;
	text: string;
}
interface Group<T> {
	kind: 'word' | 'space' | 'break';
	pieces: Piece<T>[];
}
const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

/**
 * Format-neutral horizontal flow, retaining run properties and explicit empty lines.
 * A word can span formatting runs; automatic line edges discard separating spaces.
 * Nonbreaking spaces remain inside words. The host owns font measurement and units.
 */
export function wrapStyledRuns<T extends { text: string }>(
	runs: readonly T[],
	maxWidth: number,
	measure: (content: string, run: T) => number,
	options: WrapStyledRunsOptions = {},
): T[][] {
	const groups: Group<T>[] = [];
	for (const [source, run] of runs.entries()) {
		for (const token of run.text.split(/(\r\n|[\r\n]|[ \t]+)/u).filter(Boolean)) {
			const kind = /^[\r\n]+$/u.test(token) ? 'break' : /^[ \t]+$/u.test(token) ? 'space' : 'word';
			let group = groups.at(-1);
			if (!group || group.kind !== kind || kind === 'break') {
				group = { kind, pieces: [] };
				groups.push(group);
			}
			if (kind !== 'break') group.pieces.push({ run, source, text: token });
		}
	}
	const merged = (left: Piece<T>[], right: Piece<T>[]): Piece<T>[] => {
		const result = left.map((piece) => ({ ...piece }));
		for (const piece of right) {
			const last = result.at(-1);
			if (last?.source === piece.source) last.text += piece.text;
			else result.push({ ...piece });
		}
		return result;
	};
	const width = (pieces: Piece<T>[]): number =>
		pieces.reduce((sum, piece) => {
			const advance = measure(piece.text, piece.run);
			return sum + (Number.isFinite(advance) && advance >= 0 ? advance : 0);
		}, 0);
	const result: T[][] = [];
	let line: Piece<T>[] = [];
	let spaces: Piece<T>[] = [];
	const flush = () => {
		result.push(line.map((piece) => ({ ...piece.run, text: piece.text })));
		line = [];
		spaces = [];
	};
	for (const group of groups) {
		if (group.kind === 'break') {
			flush();
			continue;
		}
		if (group.kind === 'space') {
			spaces = group.pieces;
			continue;
		}
		const separator = line.length ? spaces : [];
		const wordWidth = width(group.pieces);
		if (line.length && width(merged(line, [...separator, ...group.pieces])) > maxWidth) flush();
		else line = merged(line, separator);
		spaces = [];
		if (options.breakWords && wordWidth > maxWidth) {
			for (const piece of group.pieces) {
				for (const { segment } of graphemes.segment(piece.text)) {
					const next = { ...piece, text: segment };
					if (line.length && width(merged(line, [next])) > maxWidth) flush();
					line = merged(line, [next]);
				}
			}
		} else line = merged(line, group.pieces);
	}
	flush();
	return result;
}
