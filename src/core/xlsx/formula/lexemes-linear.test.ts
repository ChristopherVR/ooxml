import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { NUMBER_RE, readPrefix, splitSheets } from './lexemes';
import { tokenize } from './tokenizer';

/** The pattern before the linear rewrite, kept only to prove the token streams are unchanged. */
const OLD_NUMBER_RE = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;

const SAMPLES = [
	'1',
	'1.',
	'.5',
	'.',
	'..',
	'1.5',
	'1e5',
	'1E5',
	'1.5E-3',
	'.5e+2',
	'1.e3',
	'1e',
	'1e+',
	'12.34.56',
	'1..2',
	'0.0001',
	'1e5.5',
];

/** The sheet-prefix pattern before the hand-written scanner, kept to prove equivalence. */
const OLD_PREFIX_AT =
	/(\[\d+\])?([A-Za-z0-9_.\u00A1-\uFFFF]+)(?::([A-Za-z0-9_.\u00A1-\uFFFF]+))?!/y;

function oldReadPrefix(source: string, at: number): { text: string; length: number } | undefined {
	OLD_PREFIX_AT.lastIndex = at;
	const m = OLD_PREFIX_AT.exec(source);
	if (!m) return undefined;
	if (!m[1] && /^\d/.test(m[2] ?? '') && !m[3]) return undefined;
	return { text: m[0], length: m[0].length };
}

function expectSamePrefix(source: string, at: number): void {
	if (source[at] === "'") return;
	const got = readPrefix(source, at);
	const want = oldReadPrefix(source, at);
	expect(got && { text: got.prefix.text, length: got.length }, `${source}@${at}`).toEqual(want);
	if (got && want) {
		expect(got.prefix).toEqual({ text: want.text, ...splitSheets(want.text.slice(0, -1)) });
	}
}

const PREFIX_SAMPLES = [
	'Sheet1!A1',
	'Sheet1:Sheet3!A1',
	'[1]Sheet1!A1',
	'[1]Sheet1:Sheet2!B2',
	'[]Sheet1!A1',
	'[x]Sheet1!A1',
	'1A!A1',
	'[2]1A!A1',
	'1A:2B!A1',
	'Sheet1:!A1',
	':Sheet1!A1',
	'!A1',
	'a.b!A1',
	'Sheet1:Sheet2:Sheet3!A1',
	'Sheet1:Sheet2',
	'A1:B2',
	'Ünïcode!A1',
	'Sheet 1!A1',
];

function formulasOf(value: unknown, out: string[]): void {
	if (typeof value === 'string') {
		if (value.length < 2000) out.push(value);
	} else if (Array.isArray(value)) {
		for (const v of value) formulasOf(v, out);
	} else if (value && typeof value === 'object') {
		for (const v of Object.values(value)) formulasOf(v, out);
	}
}

describe('number lexeme', () => {
	it('matches exactly what the previous pattern matched', () => {
		for (const s of SAMPLES) {
			expect(NUMBER_RE.exec(s)?.[0], s).toBe(OLD_NUMBER_RE.exec(s)?.[0]);
		}
		expect(NUMBER_RE.exec('.')).toBeNull();
	});

	it('matches the old pattern on every fixture string', () => {
		const dir = new URL('./__fixtures__/', import.meta.url);
		const strings: string[] = [];
		for (const f of readdirSync(dir).filter((n) => n.endsWith('.json'))) {
			formulasOf(JSON.parse(readFileSync(new URL(f, dir), 'utf8')), strings);
		}
		expect(strings.length).toBeGreaterThan(0);
		for (const s of strings) {
			for (let i = 0; i < s.length; i++) {
				const rest = s.slice(i);
				expect(NUMBER_RE.exec(rest)?.[0]).toBe(OLD_NUMBER_RE.exec(rest)?.[0]);
			}
		}
	});

	it('reads sheet prefixes exactly as the previous pattern did', () => {
		for (const s of PREFIX_SAMPLES) {
			for (let i = 0; i < s.length; i++) expectSamePrefix(s, i);
		}
	});

	it('reads sheet prefixes like the old pattern on every fixture string', () => {
		const dir = new URL('./__fixtures__/', import.meta.url);
		const strings: string[] = [];
		for (const f of readdirSync(dir).filter((n) => n.endsWith('.json'))) {
			formulasOf(JSON.parse(readFileSync(new URL(f, dir), 'utf8')), strings);
		}
		for (const s of strings) {
			for (let i = 0; i < s.length; i++) expectSamePrefix(s, i);
		}
	});

	it('reads long runs of sheet-name characters quickly', () => {
		for (const input of [
			'a' + '.'.repeat(50_000),
			'a' + '.'.repeat(50_000) + ':',
			'a:' + '.'.repeat(50_000),
			'a.'.repeat(25_000) + ':' + 'b.'.repeat(25_000),
			'[1]' + 'a.'.repeat(25_000),
		]) {
			const started = performance.now();
			for (let i = 0; i < 50; i++) readPrefix(input, i);
			try {
				tokenize(`=${input}`);
			} catch {
				// a malformed formula may throw; only the time matters
			}
			expect(performance.now() - started).toBeLessThan(500);
		}
	});

	it('tokenizes long runs of dots and digits quickly', () => {
		for (const input of [
			'.'.repeat(50_000),
			'1' + '.'.repeat(50_000),
			'1' + '.1'.repeat(25_000),
			'1'.repeat(50_000) + 'e',
			'.'.repeat(50_000) + 'e',
		]) {
			const started = performance.now();
			try {
				tokenize(`=${input}`);
			} catch {
				// a malformed formula may throw; only the time matters
			}
			expect(performance.now() - started).toBeLessThan(500);
		}
	});
});
