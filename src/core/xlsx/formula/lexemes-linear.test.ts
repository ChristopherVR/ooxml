import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { NUMBER_RE } from './lexemes';
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
