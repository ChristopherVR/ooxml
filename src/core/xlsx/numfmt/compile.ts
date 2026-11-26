import { compileFraction, type FractionPlan, fractionSlash } from './fraction';
import { compileNumber, type NumberPlan } from './number';
import { numberFormatSections } from '../../text/number-format/sections';
import type { Condition, Section } from './types';

export interface CompiledSection extends Section {
	number?: NumberPlan;
	fraction?: FractionPlan;
}

export interface Compiled {
	/** Sections used for numbers (1-3; a lone conditional section gets a General fallback). */
	numeric: CompiledSection[];
	/** Section used for text values, if the format has one. */
	text?: CompiledSection;
	conditional: boolean;
	isDate: boolean;
}

const CACHE_LIMIT = 2000;
const cache = new Map<string, Compiled>();

function compileSection(section: CompiledSection): CompiledSection {
	if (section.kind === 'number') {
		const slash = fractionSlash(section.tokens);
		if (slash >= 0) section.fraction = compileFraction(section.tokens, slash);
		else section.number = compileNumber(section.tokens);
	}
	return section;
}

const generalSection = (): CompiledSection => ({ tokens: [{ t: 'general' }], kind: 'general' });

/** Parses a format code once; results are cached by code. */
export function compileFormat(format: string): Compiled {
	const hit = cache.get(format);
	if (hit) return hit;
	const sections = numberFormatSections(format);
	const numeric = sections.numeric.map(compileSection);
	if (numeric.length === 0) numeric.push(generalSection());
	const conditional = numeric.some((s) => s.condition !== undefined);
	if (conditional && numeric.length === 1) numeric.push(generalSection());
	const compiled: Compiled = {
		numeric,
		conditional,
		isDate: numeric.some((s) => s.kind === 'date'),
	};
	if (sections.text) compiled.text = compileSection(sections.text);
	if (cache.size >= CACHE_LIMIT) cache.clear();
	cache.set(format, compiled);
	return compiled;
}

export function matches(c: Condition, v: number): boolean {
	switch (c.op) {
		case '<':
			return v < c.value;
		case '<=':
			return v <= c.value;
		case '>':
			return v > c.value;
		case '>=':
			return v >= c.value;
		case '=':
			return v === c.value;
		default:
			return v !== c.value;
	}
}

/** A condition that only ever matches negative numbers (its section drops the minus sign). */
const onlyNegative = (c: Condition): boolean =>
	(c.op === '<' && c.value <= 0) || (c.op === '<=' && c.value < 0) || (c.op === '=' && c.value < 0);

export interface Choice {
	section: CompiledSection;
	/** Whether the section shows the absolute value (no minus sign). */
	abs: boolean;
}

/** Picks the section for a number the way Excel does (`undefined`: no section applies, `####`). */
export function chooseSection(c: Compiled, v: number): Choice | undefined {
	const [s1, s2, s3] = c.numeric;
	if (!s1) return undefined;
	if (!c.conditional) {
		if (v > 0 || !s2) return { section: s1, abs: false };
		if (v < 0) return { section: s2, abs: true };
		return { section: s3 ?? s1, abs: false };
	}
	const c1 = s1.condition;
	const c2 = s2?.condition;
	if (!c1 || matches(c1, v)) return { section: s1, abs: c1 ? onlyNegative(c1) : false };
	if (!s2) return undefined;
	if (c2) {
		if (matches(c2, v)) return { section: s2, abs: onlyNegative(c2) };
		return s3 ? { section: s3, abs: false } : undefined;
	}
	if (s3) return v < 0 ? { section: s2, abs: true } : { section: s3, abs: false };
	return { section: s2, abs: v < 0 && c1.op !== '=' && c1.value <= 0 };
}
