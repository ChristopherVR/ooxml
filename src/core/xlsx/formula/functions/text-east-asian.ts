import { ERR, isError, RefValue } from '../values';
import { scalar, spec, str } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Text';

/** Half-width Katakana (U+FF61..U+FF9F) paired with their full-width equivalents. */
const HALF_TO_FULL_KANA =
	'。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜';

const DAKUTEN_PAIRS: ReadonlyArray<readonly [half: string, full: string]> = [
	['ｳ', 'ヴ'],
	['ｶ', 'ガ'],
	['ｷ', 'ギ'],
	['ｸ', 'グ'],
	['ｹ', 'ゲ'],
	['ｺ', 'ゴ'],
	['ｻ', 'ザ'],
	['ｼ', 'ジ'],
	['ｽ', 'ズ'],
	['ｾ', 'ゼ'],
	['ｿ', 'ゾ'],
	['ﾀ', 'ダ'],
	['ﾁ', 'ヂ'],
	['ﾂ', 'ヅ'],
	['ﾃ', 'デ'],
	['ﾄ', 'ド'],
	['ﾊ', 'バ'],
	['ﾋ', 'ビ'],
	['ﾌ', 'ブ'],
	['ﾍ', 'ベ'],
	['ﾎ', 'ボ'],
	['ﾜ', 'ヷ'],
	['ｦ', 'ヺ'],
];

const HANDAKUTEN_PAIRS: ReadonlyArray<readonly [half: string, full: string]> = [
	['ﾊ', 'パ'],
	['ﾋ', 'ピ'],
	['ﾌ', 'プ'],
	['ﾍ', 'ペ'],
	['ﾎ', 'ポ'],
];

const FULL_TO_HALF = new Map<string, string>([
	['\u3000', ' '],
	['\uFFE5', '\u00A5'],
]);
const HALF_DAKUTEN = new Map<string, string>();
const HALF_HANDAKUTEN = new Map<string, string>();

for (let i = 0; i < HALF_TO_FULL_KANA.length; i++) {
	const full = HALF_TO_FULL_KANA[i]!;
	const half = String.fromCharCode(0xff61 + i);
	FULL_TO_HALF.set(full, half);
}
for (const [half, full] of DAKUTEN_PAIRS) {
	FULL_TO_HALF.set(full, `${half}\uFF9E`);
	HALF_DAKUTEN.set(half, full);
}
for (const [half, full] of HANDAKUTEN_PAIRS) {
	FULL_TO_HALF.set(full, `${half}\uFF9F`);
	HALF_HANDAKUTEN.set(half, full);
}

/** Converts full-width ASCII and Katakana characters to half-width equivalents. */
export function toHalfWidth(text: string): string {
	let out = '';
	for (let i = 0; i < text.length; i++) {
		const ch = text[i]!;
		const code = ch.charCodeAt(0);
		if (code >= 0xff01 && code <= 0xff5e) {
			out += String.fromCharCode(code - 0xfee0);
			continue;
		}
		out += FULL_TO_HALF.get(ch) ?? ch;
	}
	return out;
}

/** Converts half-width ASCII and Katakana characters to full-width equivalents. */
export function toFullWidth(text: string): string {
	let out = '';
	for (let i = 0; i < text.length; i++) {
		const ch = text[i]!;
		const code = ch.charCodeAt(0);
		if (code === 0x0020) {
			out += '\u3000';
			continue;
		}
		if (code >= 0x0021 && code <= 0x007e) {
			out += String.fromCharCode(code + 0xfee0);
			continue;
		}
		if (code === 0x00a5) {
			out += '\uFFE5';
			continue;
		}
		if (code >= 0xff61 && code <= 0xff9f) {
			const next = text[i + 1];
			if (next === '\uFF9E') {
				const voiced = HALF_DAKUTEN.get(ch);
				if (voiced) {
					out += voiced;
					i++;
					continue;
				}
			} else if (next === '\uFF9F') {
				const semi = HALF_HANDAKUTEN.get(ch);
				if (semi) {
					out += semi;
					i++;
					continue;
				}
			}
			out += HALF_TO_FULL_KANA[code - 0xff61] ?? ch;
			continue;
		}
		out += ch;
	}
	return out;
}

const fullWidthSpec = (name: 'DBCS' | 'JIS'): FunctionSpec =>
	spec(
		name,
		C,
		`${name}(text)`,
		'Converts half-width (single-byte) letters and Katakana to full-width (double-byte) characters.',
		1,
		1,
		(args) => {
			const v = scalar(args[0]);
			if (isError(v)) return v;
			return toFullWidth(str(v));
		},
	);

export const TEXT_EAST_ASIAN_FUNCTIONS: FunctionSpec[] = [
	spec(
		'ASC',
		C,
		'ASC(text)',
		'Converts full-width (double-byte) letters and Katakana to half-width (single-byte) characters.',
		1,
		1,
		(args) => {
			const v = scalar(args[0]);
			if (isError(v)) return v;
			return toHalfWidth(str(v));
		},
	),
	fullWidthSpec('DBCS'),
	fullWidthSpec('JIS'),
	spec(
		'PHONETIC',
		C,
		'PHONETIC(reference)',
		'Extracts the phonetic (furigana) characters from a text string reference.',
		1,
		1,
		(args, ctx) => {
			const ref = args[0];
			if (!(ref instanceof RefValue)) return isError(ref ?? null) ? scalar(ref) : ERR.VALUE;
			if (ref.areas.length === 0) return ERR.REF;
			if (ref.areas.length > 1) return ERR.NA;
			const area = ref.areas[0]!;
			const { row, col } = area.range.start;
			const cell = ctx.workbook.sheets[area.sheet]?.rows.get(row)?.get(col);
			if (cell?.phonetic?.runs.length) return cell.phonetic.runs.map((run) => run.text).join('');
			const value = ctx.readCell(area.sheet, row, col);
			if (isError(value)) return value;
			return typeof value === 'string' ? value : '';
		},
		['any'],
	),
];
