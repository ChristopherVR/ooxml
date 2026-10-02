import { describe, expect, it } from 'vitest';
import {
	inspectVisioEmfAdmission as inspect,
	type VisioEmfAdmissionOptions,
} from './emf-admission.js';
import { emf, record, setWord } from './emf-admission.test-fixtures.js';
const rect = record(43, [0, 0, 10, 10]);
const mapping = [record(17, [8]), record(9, [1, 1]), record(11, [2, 2])];
const codes = (bytes: Uint8Array) => inspect(bytes).diagnostics.map((d) => d.code);

// Fail-closed regressions from independent synthetic converter3.5.1 review.
// They assert admission only, never import or execute the converter in this core test suite.
describe('known converter mapping and arc compatibility exclusions', () => {
	it('rejects anisotropic to MM_TEXT transitions that retain stale converter scale', () => {
		expect(codes(emf([...mapping, record(17, [1]), rect]))).toContain(
			'mapping-transition-fidelity',
		);
		expect(inspect(emf([record(17, [1]), rect])).status).toBe('admitted');
	});
	it.each([
		record(9, [2, 2]),
		record(10, [10, 10]),
		record(11, [4, 4]),
		record(12, [10, 10]),
		record(17, [1]),
	])('rejects any mapping mutation inside saved state %#', (changed) => {
		expect(codes(emf([...mapping, record(33), changed, record(34, [-1]), rect]))).toContain(
			'mapping-restore-fidelity',
		);
	});
	it('allows balanced save/restore without mapping edits', () => {
		expect(inspect(emf([...mapping, record(33), rect, record(34, [-1]), rect])).status).toBe(
			'admitted',
		);
	});
	it.each([{ setup: [] }, { setup: [record(9, [1, 1])] }, { setup: [record(11, [2, 2])] }])(
		'rejects partial anisotropic defaults before drawing %#',
		({ setup }) => {
			expect(codes(emf([record(17, [8]), ...setup, rect]))).toContain('mapping-incomplete');
		},
	);
	it('also requires complete mapping before clipping or current-position capture', () => {
		for (const action of [record(30, [0, 0, 10, 10]), record(27, [10, 10])])
			expect(
				codes(emf([record(17, [8]), record(11, [2, 2]), action, record(9, [1, 1]), rect])),
			).toContain('mapping-incomplete');
	});
	it('requires an explicit window origin for nonzero header origins', () => {
		const incomplete = setWord(setWord(emf([...mapping, rect]), 8, 10), 12, 10);
		expect(codes(incomplete)).toContain('mapping-incomplete');
		const complete = setWord(setWord(emf([...mapping, record(10, [0, 0]), rect]), 8, 10), 12, 10);
		expect(inspect(complete).status).toBe('admitted');
	});
	it.each([45, 47])('rejects same-ray full-ellipse semantics for record%i', (type) => {
		for (const endpoints of [
			[40, 10, 40, 10],
			[40, 10, 60, 10],
			[0, 0, -20, -10],
		])
			expect(codes(emf([record(type, [0, 0, 40, 20, ...endpoints])]))).toContain(
				'arc-full-ellipse-fidelity',
			);
		// Opposite radial directions form a half ellipse and do not trigger this exclusion.
		expect(inspect(emf([record(type, [0, 0, 40, 20, 40, 10, 0, 10])])).status).toBe('admitted');
	});
	it('uses exact doubled-center arithmetic with odd bounds and bounded large coordinates', () => {
		expect(codes(emf([record(45, [0, 0, 41, 21, 41, 21, 82, 42])]))).toContain(
			'arc-full-ellipse-fidelity',
		);
		expect(
			codes(
				emf([record(45, [-1000000, -1000000, 1000000, 1000000, 1000000, 500000, 500000, 250000])]),
			),
		).toContain('arc-full-ellipse-fidelity');
	});
});

describe('intrinsic EMF input and data-only options boundary', () => {
	it.each(['buffer', 'byteOffset', 'byteLength'])(
		'never invokes a source-owned %s accessor',
		(key) => {
			const bytes = emf([rect]);
			let calls = 0;
			Object.defineProperty(bytes, key, {
				get() {
					calls++;
					throw new Error('do not expose source exceptions');
				},
			});
			expect(inspect(bytes).status).toBe('admitted');
			expect(calls).toBe(0);
		},
	);
	it('uses the intrinsic typed-array tag and rejects proxy/fake typed arrays', () => {
		const bytes = emf([rect]);
		Object.defineProperty(bytes, Symbol.toStringTag, {
			get() {
				throw new Error('tag accessor');
			},
		});
		expect(inspect(bytes).status).toBe('admitted');
		for (const fake of [
			new Proxy(bytes, {}),
			Object.create(Uint8Array.prototype),
			new Uint16Array(100),
		])
			expect(inspect(fake as Uint8Array).status).toBe('invalid');
	});
	it.each(['maxDiagnostics', 'maxRecords', 'maxCoordinate'])(
		'rejects own %s accessors without executing them',
		(key) => {
			const bytes = emf([rect]);
			let calls = 0;
			const options = Object.defineProperty({}, key, {
				get() {
					calls++;
					structuredClone(bytes, { transfer: [bytes.buffer] });
					return 1;
				},
			});
			const result = inspect(bytes, options);
			expect(result.status).toBe('invalid');
			expect(result.diagnostics[0]?.code).toBe('limits');
			expect(calls).toBe(0);
		},
	);
	it('rejects inherited options without invoking their accessors', () => {
		let calls = 0;
		const options = Object.create({
			get maxRecords() {
				calls++;
				return 2;
			},
		}) as VisioEmfAdmissionOptions;
		expect(inspect(emf([rect]), options).status).toBe('invalid');
		expect(calls).toBe(0);
	});
	it('sanitizes throwing option-proxy traps and reads options before the input buffer', () => {
		const throws = new Proxy(
			{},
			{
				getOwnPropertyDescriptor() {
					throw new Error('private detail');
				},
			},
		);
		const failure = inspect(emf([rect]), throws);
		expect(failure.status).toBe('invalid');
		expect(JSON.stringify(failure)).not.toContain('private detail');
		const bytes = emf([rect]);
		let detached = false;
		const detaches = new Proxy(
			{},
			{
				getOwnPropertyDescriptor() {
					if (!detached) {
						structuredClone(bytes, { transfer: [bytes.buffer] });
						detached = true;
					}
					return undefined;
				},
			},
		);
		expect(inspect(bytes, detaches).diagnostics[0]?.code).toBe('input');
	});
});
