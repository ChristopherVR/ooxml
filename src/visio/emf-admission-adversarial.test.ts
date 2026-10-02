import { describe, expect, it } from 'vitest';
import {
	inspectVisioEmfAdmission as inspect,
	type VisioEmfAdmissionOptions,
} from './emf-admission.js';
import {
	checksumComment,
	emf,
	polygon16,
	record,
	setWord,
	sourceComment,
	triangle,
} from './emf-admission.test-fixtures.js';
const has = (b: Uint8Array, code: string) =>
	expect(inspect(b).diagnostics.map((d) => d.code)).toContain(code);

describe('EMF malformed and adversarial inputs', () => {
	it.each([
		null,
		undefined,
		'',
		{},
		new Uint16Array(40),
		new Uint8Array(new SharedArrayBuffer(100)),
	])('rejects invalid runtime input %# without throwing', (b) => {
		expect(inspect(b as Uint8Array).status).toBe('invalid');
	});
	it('rejects detached and resizable backing buffers', () => {
		const b = emf([]);
		structuredClone(b, { transfer: [b.buffer] });
		expect(inspect(b).status).toBe('invalid');
		const ctor = ArrayBuffer as unknown as new (
			n: number,
			options: { maxByteLength: number },
		) => ArrayBuffer;
		const resizable = new Uint8Array(new ctor(100, { maxByteLength: 200 }));
		expect(inspect(resizable).status).toBe('invalid');
	});
	it('does not read past any truncated generated stream or record', () => {
		const b = checksumComment(
			emf([
				sourceComment(),
				record(17, [8]),
				record(11, [10, 10]),
				polygon16(triangle),
				record(75, [0, 5]),
			]),
		);
		for (let end = 0; end < b.length; end++)
			expect(inspect(b.subarray(0, end)).status).not.toBe('admitted');
		for (const type of [
			9, 17, 27, 33, 34, 37, 38, 39, 40, 42, 43, 45, 47, 48, 54, 70, 75, 86, 87,
		]) {
			const payload = record(type, Array(12).fill(0));
			for (let size = 8; size < payload.length; size += 4) {
				const short = setWord(payload.subarray(0, size), 4, size);
				expect(() => inspect(emf([short]))).not.toThrow();
			}
		}
	});
	it('rejects duplicate/truncated headers, missing EOF and trailing data', () => {
		has(emf([record(1)]), 'header');
		has(setWord(emf([]), 108, 33), 'payload-size');
		const b = emf([record(14, [0, 16, 20]), record(33)]);
		has(b, 'trailing-data');
		has(setWord(emf([]), 52, 3), 'record-count');
	});
	it('rejects incorrect fixed sizes and record-local variable ranges', () => {
		has(emf([record(38, [1, 0, 1, 0])]), 'payload-size');
		has(emf([setWord(polygon16(triangle), 24, 0xffffffff)]), 'payload-size');
		has(emf([record(70, [500, 0x43494447])]), 'comment-range');
		has(emf([record(75, [100, 5])]), 'region-data');
		has(emf([record(75, [32, 5, 32, 1, 0xffffffff, 0, 0, 0, 1, 1])]), 'region-data');
	});
	it('rejects unsafe header descriptions and extension payloads', () => {
		let b = setWord(emf([]), 60, 100);
		b = setWord(b, 64, 100);
		has(b, 'header');
		has(setWord(emf([]), 88, 20), 'header');
		has(setWord(emf([]), 58, 1), 'header');
	});
	it('rejects invalid handles, duplicate live objects, deletion and saved-state selection', () => {
		for (const handle of [0, 32, 0x80000001])
			has(emf([record(38, [handle, 0, 1, 0, 0])]), 'object-handle');
		has(emf([record(38, [1, 0, 1, 0, 0]), record(38, [1, 0, 1, 0, 0])]), 'object-handle');
		has(emf([record(40, [1])]), 'object-reference');
		has(
			emf([
				record(38, [1, 0, 1, 0, 0]),
				record(37, [1]),
				record(33),
				record(37, [0x80000007]),
				record(40, [1]),
				record(34, [-1]),
			]),
			'selected-object',
		);
	});
	it('does not invent an invalid handle/restore diagnosis after unknown state-changing records', () => {
		const r = inspect(emf([record(82), record(37, [1]), record(40, [1]), record(34, [-1])]));
		expect(r.status).toBe('unsupported');
		expect(r.diagnostics.every((d) => d.kind === 'unsupported')).toBe(true);
	});
	it('rejects invalid restores and unbalanced supported stacks', () => {
		for (const n of [-2, 0, 1]) has(emf([record(33), record(34, [n])]), 'restore-state');
		has(emf([record(33)]), 'unbalanced-state');
		expect(inspect(emf([record(33), record(33), record(34, [-2])])).status).toBe('admitted');
	});
	it('rejects degenerate arcs and unsupported styles rather than normalizing them', () => {
		has(emf([record(45, [0, 0, 10, 10, 5, 5, 10, 10])]), 'parameter');
		for (const style of [1, 2, 3, 4, 6, 7, 8, 0x10000])
			has(emf([record(38, [1, style, 1, 0, 0])]), 'pen-style');
		has(emf([record(39, [1, 2, 0, 0])]), 'brush-style');
	});
	it('rejects nonpositive, extreme and highly anisotropic mapping', () => {
		for (const size of [
			[0, 1],
			[-1, 1],
		])
			has(emf([record(17, [8]), record(9, size)]), 'parameter');
		has(emf([record(17, [8]), record(11, [4097, 1])]), 'mapping-scale');
		has(emf([record(17, [8]), record(11, [4096, 1]), record(9, [1, 4096])]), 'mapping-scale');
		has(emf([record(17, [8]), record(11, [100, 100]), record(27, [10001, 0])]), 'coordinate-limit');
	});
	it('charges repeated object creation even when handles are reused', () => {
		const records = Array.from({ length: 10 }, () => [
			record(38, [1, 0, 1, 0, 0]),
			record(40, [1]),
		]).flat();
		expect(inspect(emf(records), { maxObjects: 2 }).status).toBe('budget-exceeded');
	});
	it.each<[string, Uint8Array, VisioEmfAdmissionOptions]>([
		['input', emf([]), { maxInputBytes: 100 }],
		['records', emf([record(33), record(34, [-1])]), { maxRecords: 3 }],
		['types', emf([record(43, [0, 0, 10, 10])]), { maxRecordTypes: 1 }],
		['handles', emf([]), { maxHandles: 10 }],
		['points', emf([polygon16(triangle), polygon16(triangle)]), { maxPoints: 4 }],
		[
			'regions',
			emf([record(75, [64, 5, 32, 1, 2, 32, 0, 0, 20, 20, 0, 0, 10, 10, 10, 10, 20, 20])]),
			{ maxRegionRects: 1 },
		],
		['state', emf([record(33), record(33)]), { maxStateDepth: 1 }],
		['clip ops', emf([record(75, [0, 5]), record(75, [0, 5])]), { maxClipOperations: 1 }],
		[
			'clip depth',
			emf([record(30, [0, 0, 10, 10]), record(30, [0, 0, 9, 9])]),
			{ maxClipDepth: 1 },
		],
		['comments', checksumComment(emf([sourceComment()])), { maxCommentBytes: 1 }],
		['allocation', emf([polygon16(triangle)]), { maxAllocationUnits: 1 }],
		['pen', emf([record(38, [1, 0, 3, 0, 0])]), { maxPenWidth: 2 }],
	])('enforces aggregate %s before variable allocation', (_, b, options) => {
		expect(inspect(b, options).status).toBe('budget-exceeded');
	});
	it('charges clip ancestry multiplied by draw work, not records alone', () => {
		const one = inspect(emf([polygon16(triangle)]));
		const clips = Array.from({ length: 8 }, () => record(30, [0, 0, 10, 10]));
		const many = inspect(emf([...clips, ...Array.from({ length: 10 }, () => polygon16(triangle))]));
		expect(many.metrics.allocationUnits).toBeGreaterThan(one.metrics.allocationUnits * 50);
	});
	it('validates EOF palette sizes/counts and caps palette allocation', () => {
		has(setWord(emf([]), 108 + 8, 1), 'eof');
		expect(inspect(setWord(emf([]), 68, 257)).status).toBe('budget-exceeded');
	});
	it('never throws or returns unbounded diagnostics on deterministic byte mutations', () => {
		const source = checksumComment(
			emf([sourceComment(), record(38, [1, 0, 1, 0, 0]), record(37, [1]), polygon16(triangle)]),
		);
		let seed = 42;
		for (let i = 0; i < 1500; i++) {
			seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
			const b = source.slice(),
				at = seed % b.length;
			b[at] = (b[at] ?? 0) ^ (1 << (seed % 8));
			const r = inspect(b, { maxDiagnostics: 4 });
			expect(r.diagnostics.length).toBeLessThanOrEqual(4);
			expect(r.renderingEnabled).toBe(false);
		}
	});
});
