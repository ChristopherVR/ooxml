import type { Condition, ConditionOp, DatePart, Section, SectionKind, Token } from './types.js';

const NAMED_COLORS: Readonly<Record<string, string>> = {
	black: '#000000',
	white: '#FFFFFF',
	red: '#FF0000',
	green: '#00FF00',
	blue: '#0000FF',
	yellow: '#FFFF00',
	magenta: '#FF00FF',
	cyan: '#00FFFF',
};

/** The default 56-colour palette `[Color1]`..`[Color56]` refer to. */
const PALETTE_56 = [
	'000000',
	'FFFFFF',
	'FF0000',
	'00FF00',
	'0000FF',
	'FFFF00',
	'FF00FF',
	'00FFFF',
	'800000',
	'008000',
	'000080',
	'808000',
	'800080',
	'008080',
	'C0C0C0',
	'808080',
	'9999FF',
	'993366',
	'FFFFCC',
	'CCFFFF',
	'660066',
	'FF8080',
	'0066CC',
	'CCCCFF',
	'000080',
	'FF00FF',
	'FFFF00',
	'00FFFF',
	'800080',
	'800000',
	'008080',
	'0000FF',
	'00CCFF',
	'CCFFFF',
	'CCFFCC',
	'FFFF99',
	'99CCFF',
	'FF99CC',
	'CC99FF',
	'FFCC99',
	'3366FF',
	'33CCCC',
	'99CC00',
	'FFCC00',
	'FF9900',
	'FF6600',
	'666699',
	'969696',
	'003366',
	'339966',
	'003300',
	'333300',
	'993300',
	'993366',
	'333399',
	'333333',
];

/** Splits a format code into its `;` separated sections (quotes, escapes and brackets respected). */
export function splitSections(format: string): string[] {
	const sections: string[] = [];
	let current = '';
	for (let i = 0; i < format.length; i++) {
		const ch = format.charAt(i);
		if (ch === '"') {
			const end = format.indexOf('"', i + 1);
			const stop = end < 0 ? format.length : end + 1;
			current += format.slice(i, stop);
			i = stop - 1;
		} else if (ch === '\\' || ch === '_' || ch === '*') {
			current += format.slice(i, i + 2);
			i++;
		} else if (ch === '[') {
			const end = format.indexOf(']', i + 1);
			const stop = end < 0 ? format.length : end + 1;
			current += format.slice(i, stop);
			i = stop - 1;
		} else if (ch === ';') {
			sections.push(current);
			current = '';
		} else {
			current += ch;
		}
	}
	sections.push(current);
	return sections;
}

interface BracketResult {
	color?: string;
	condition?: Condition;
	token?: Token;
}

function parseBracket(content: string): BracketResult {
	const lower = content.toLowerCase();
	const named = NAMED_COLORS[lower];
	if (named) return { color: named };
	const indexed = /^color\s*(\d+)$/i.exec(content);
	if (indexed) {
		const hex = PALETTE_56[Number(indexed[1]) - 1];
		return hex ? { color: `#${hex}` } : {};
	}
	const cond = /^(<=|>=|<>|<|>|=)\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)$/i.exec(content);
	if (cond) return { condition: { op: cond[1] as ConditionOp, value: Number(cond[2]) } };
	if (/^(h+|m+|s+)$/i.test(content)) {
		const unit = lower.charAt(0) as 'h' | 'm' | 's';
		return { token: { t: 'elapsed', unit, width: content.length } };
	}
	if (content.startsWith('$')) {
		const dash = content.indexOf('-', 1);
		const symbol = dash < 0 ? content.slice(1) : content.slice(1, dash);
		return symbol ? { token: { t: 'lit', v: symbol } } : {};
	}
	return {};
}

function runLength(s: string, i: number, lowerCh: string): number {
	let n = 0;
	while (i + n < s.length && s.charAt(i + n).toLowerCase() === lowerCh) n++;
	return n;
}

function datePartFor(ch: string, n: number): DatePart {
	switch (ch) {
		case 'y':
			return n <= 2 ? 'y' : 'yyyy';
		case 'b':
			return n <= 2 ? 'b' : 'bbbb';
		case 'm':
			return (['m', 'mm', 'mmm', 'mmmm', 'mmmmm'] as const)[Math.min(n, 6) - 1] ?? 'mmmm';
		case 'd':
			return (['d', 'dd', 'ddd', 'dddd'] as const)[Math.min(n, 4) - 1] ?? 'dddd';
		case 'h':
			return n === 1 ? 'h' : 'hh';
		default:
			return n === 1 ? 's' : 'ss';
	}
}

/** Tokenizes one section and extracts its colour and condition. */
export function tokenizeSection(src: string): Section {
	const tokens: Token[] = [];
	let color: string | undefined;
	let condition: Condition | undefined;
	const lit = (v: string): void => {
		const last = tokens[tokens.length - 1];
		if (last?.t === 'lit') last.v += v;
		else tokens.push({ t: 'lit', v });
	};
	for (let i = 0; i < src.length; i++) {
		const ch = src.charAt(i);
		const lower = ch.toLowerCase();
		if (ch === '"') {
			const end = src.indexOf('"', i + 1);
			const stop = end < 0 ? src.length : end;
			lit(src.slice(i + 1, stop));
			i = stop;
		} else if (ch === '\\') {
			if (i + 1 < src.length) lit(src.charAt(i + 1));
			i++;
		} else if (ch === '_') {
			lit(' ');
			i++;
		} else if (ch === '*') {
			i++;
		} else if (ch === '[') {
			const end = src.indexOf(']', i + 1);
			const stop = end < 0 ? src.length : end;
			const res = parseBracket(src.slice(i + 1, stop));
			if (res.color) color = res.color;
			if (res.condition) condition = res.condition;
			if (res.token?.t === 'lit') lit(res.token.v);
			else if (res.token) tokens.push(res.token);
			i = stop;
		} else if (ch === '0' || ch === '#' || ch === '?') {
			tokens.push({ t: 'digit', c: ch });
		} else if (ch >= '1' && ch <= '9') {
			tokens.push({ t: 'num', v: ch });
		} else if (ch === '.') {
			tokens.push({ t: 'point' });
		} else if (ch === ',') {
			tokens.push({ t: 'comma' });
		} else if (ch === '%') {
			tokens.push({ t: 'percent' });
		} else if (ch === '/') {
			tokens.push({ t: 'slash' });
		} else if (ch === '@') {
			tokens.push({ t: 'text' });
		} else if (lower === 'e' && (src.charAt(i + 1) === '+' || src.charAt(i + 1) === '-')) {
			tokens.push({ t: 'exp', plus: src.charAt(i + 1) === '+' });
			i++;
		} else if (lower === 'g' && src.slice(i, i + 7).toLowerCase() === 'general') {
			tokens.push({ t: 'general' });
			i += 6;
		} else if (lower === 'a' && src.slice(i, i + 5).toUpperCase() === 'AM/PM') {
			tokens.push({ t: 'ampm', kind: 'AM/PM' });
			i += 4;
		} else if (lower === 'a' && src.slice(i, i + 3).toUpperCase() === 'A/P') {
			tokens.push({ t: 'ampm', kind: ch === 'A' ? 'A/P' : 'a/p' });
			i += 2;
		} else if (lower === 'e' || lower === 'g') {
			const n = runLength(src, i, lower);
			tokens.push({ t: 'date', part: lower });
			i += n - 1;
		} else if ('ymdhsb'.includes(lower)) {
			const n = runLength(src, i, lower);
			tokens.push({ t: 'date', part: datePartFor(lower, n) });
			i += n - 1;
		} else {
			lit(ch);
		}
	}
	const kind = classify(tokens);
	const section: Section = { tokens: kind === 'date' ? normalizeDateTokens(tokens) : tokens, kind };
	if (color) section.color = color;
	if (condition) section.condition = condition;
	return section;
}

function classify(tokens: readonly Token[]): SectionKind {
	if (tokens.some((t) => t.t === 'text')) return 'text';
	if (tokens.some((t) => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm')) return 'date';
	if (tokens.some((t) => t.t === 'general')) return 'general';
	return 'number';
}

const isHourToken = (t: Token | undefined): boolean =>
	(t?.t === 'date' && (t.part === 'h' || t.part === 'hh')) ||
	(t?.t === 'elapsed' && t.unit === 'h');
const isSecondToken = (t: Token | undefined): boolean =>
	(t?.t === 'date' && (t.part === 's' || t.part === 'ss')) ||
	(t?.t === 'elapsed' && t.unit === 's');

/** Resolves month/minute ambiguity, sub-second digits, and turns number-only tokens into literals. */
function normalizeDateTokens(tokens: readonly Token[]): Token[] {
	const out: Token[] = [];
	const lit = (v: string): void => {
		const last = out[out.length - 1];
		if (last?.t === 'lit') last.v += v;
		else out.push({ t: 'lit', v });
	};
	for (let i = 0; i < tokens.length; i++) {
		const tok = tokens[i];
		if (!tok) continue;
		if (tok.t === 'point') {
			let n = 0;
			while (tokens[i + 1 + n]?.t === 'digit' && (tokens[i + 1 + n] as { c: string }).c === '0')
				n++;
			if (n > 0) {
				out.push({ t: 'subsec', digits: n });
				i += n;
			} else lit('.');
		} else if (tok.t === 'digit') lit(tok.c);
		else if (tok.t === 'num') lit(tok.v);
		else if (tok.t === 'comma') lit(',');
		else if (tok.t === 'percent') lit('%');
		else if (tok.t === 'slash') lit('/');
		else if (tok.t === 'exp') lit(tok.plus ? 'E+' : 'E-');
		else if (tok.t === 'general') lit('General');
		else out.push(tok.t === 'lit' ? { t: 'lit', v: tok.v } : tok);
	}
	const timeTokens = out.filter((t) => t.t === 'date' || t.t === 'elapsed');
	for (let k = 0; k < timeTokens.length; k++) {
		const tok = timeTokens[k];
		if (tok?.t !== 'date' || (tok.part !== 'm' && tok.part !== 'mm')) continue;
		if (isHourToken(timeTokens[k - 1]) || isSecondToken(timeTokens[k + 1])) {
			tok.part = tok.part === 'm' ? 'min' : 'mmin';
		}
	}
	return out;
}
