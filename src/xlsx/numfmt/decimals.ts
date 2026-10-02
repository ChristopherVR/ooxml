// Excel's Increase Decimal / Decrease Decimal commands on a number format code.
import type { CellValue } from '../model.js';
import { formatGeneral } from './general.js';
import { splitSections, tokenizeSection } from './tokenizer.js';

/** The fixed format Excel switches a General cell to, sized from what it shows now. */
function fromGeneral(delta: 1 | -1, sample: CellValue | undefined): string {
	const text = typeof sample === 'number' ? formatGeneral(Math.abs(sample)) : '';
	const sci = /^(\d)(?:\.(\d+))?E/.exec(text);
	const shown = sci ? (sci[2]?.length ?? 0) : (/\.(\d+)/.exec(text)?.[1]?.length ?? 0);
	const next = Math.max(0, shown + delta);
	const body = next ? `0.${'0'.repeat(next)}` : '0';
	return sci ? `${body}E+00` : body;
}

/** Positions of the number part: the last digit placeholder and the decimal point before it. */
function scan(section: string): { lastDigit: number; dot: number } {
	let quoted = false;
	let bracket = false;
	let lastDigit = -1;
	let dot = -1;
	for (let i = 0; i < section.length; i++) {
		const ch = section[i] ?? '';
		if (quoted) {
			if (ch === '"') quoted = false;
			continue;
		}
		if (bracket) {
			if (ch === ']') bracket = false;
			continue;
		}
		if (ch === '\\' || ch === '_' || ch === '*') i++;
		else if (ch === '"') quoted = true;
		else if (ch === '[') bracket = true;
		else if (ch === '0' || ch === '#' || ch === '?') lastDigit = i;
		else if (ch === '.' && dot < 0) dot = i;
		else if ((ch === 'e' || ch === 'E') && lastDigit >= 0) break;
	}
	return { lastDigit, dot: dot >= 0 && dot < lastDigit ? dot : -1 };
}

function stepSection(section: string, delta: 1 | -1): string {
	const kind = tokenizeSection(section).kind;
	if (kind !== 'number') return section;
	if (/\?\s*\/|\/\s*[?#0-9]/.test(section.replace(/"[^"]*"/g, ''))) return section;
	const { lastDigit, dot } = scan(section);
	if (lastDigit < 0) return section;
	const head = section.slice(0, lastDigit + 1);
	const tail = section.slice(lastDigit + 1);
	if (delta > 0) return dot >= 0 ? `${head}0${tail}` : `${head}.0${tail}`;
	if (dot < 0) return section;
	if (lastDigit - dot <= 1) return section.slice(0, dot) + tail;
	return section.slice(0, lastDigit) + tail;
}

/**
 * The format after Excel's Increase Decimal (`delta` 1) or Decrease Decimal (-1). Every numeric
 * section gains or loses one decimal place (scientific formats step the mantissa); quoted text,
 * escapes, colours and conditions are kept. `General` becomes a fixed format sized from the
 * decimals `sampleValue` currently shows (`1.5` General +1 gives `0.00`). Text (`@`), date,
 * time and fraction sections are unchanged, and decreasing past zero decimals is a no-op.
 */
export function stepDecimals(format: string, delta: 1 | -1, sampleValue?: CellValue): string {
	const code = format.trim() === '' ? 'General' : format;
	if (/^general$/i.test(code.trim())) return fromGeneral(delta, sampleValue);
	return splitSections(code)
		.map((section) => stepSection(section, delta))
		.join(';');
}
