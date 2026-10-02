import { describe, expect, it } from 'vitest';
import {
	inspectVisioEmfAdmission as inspect,
	VISIO_EMF_ADMISSION_LIMITS,
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

const codes = (bytes: Uint8Array) => inspect(bytes).diagnostics.map((d) => d.code);
const map = [
	record(17, [8]),
	record(11, [100, 100]),
	record(9, [200, 200]),
	record(10, [-100, -100]),
];

describe('bounded classic EMF admission, rendering disabled', () => {
	it.each([88, 100, 108])(
		'admits the generated fixed %i-byte header and literal rectangle subset',
		(size) => {
			const result = inspect(emf([record(43, [0, 0, 90, 90])], size));
			expect(result).toMatchObject({
				status: 'admitted',
				renderingEnabled: false,
				scanComplete: true,
				diagnostics: [],
				header: { pixelWidth: 100, pixelHeight: 100 },
			});
		},
	);
	it('checks a nonzero-offset view without retaining or changing the source', () => {
		const source = emf([record(42, [0, 0, 99, 99])]),
			b = new Uint8Array(source.length + 20).fill(255);
		b.set(source, 8);
		const slice = b.subarray(8, 8 + source.length);
		const result = inspect(slice);
		expect(result.status).toBe('admitted');
		expect(slice).toEqual(source);
		expect(JSON.stringify(result)).not.toContain('buffer');
	});
	it('bounds restored mapping but rejects converter-incompatible origin restoration', () => {
		const result = inspect(
			emf([
				...map,
				record(33),
				record(10, [-200, -200]),
				polygon16(triangle),
				record(34, [-1]),
				polygon16([
					[-100, -100],
					[0, -100],
					[-100, 0],
				]),
			]),
		);
		expect(result.status).toBe('unsupported');
		expect(result.diagnostics.map((d) => d.code)).toContain('mapping-restore-fidelity');
		expect(result.metrics.maxMappedCoordinate).toBe(105);
		expect(result.metrics.points).toBe(6);
		expect(result.metrics.peakStateDepth).toBe(1);
	});
	it('validates ARC45 and PIE47 radial points outside their bounding rectangle', () => {
		const result = inspect(
			emf([
				record(45, [0, 0, 100, 100, 200, 50, 50, -100]),
				record(47, [0, 0, 100, 100, 200, 50, 50, -100]),
			]),
		);
		expect(result.status).toBe('admitted');
		expect(result.recordTypes.map((r) => r.type)).toContain(47);
		expect(codes(emf([record(55, [0, 0, 100, 100, 200, 50, 50, -100])]))).toContain(
			'unsupported-record',
		);
	});
	it('tracks current position for LineTo and restores it with the DC', () => {
		const result = inspect(
			emf([
				...map,
				record(27, [-100, -100]),
				record(33),
				record(27, [0, 0]),
				record(34, [-1]),
				record(54, [-90, -90]),
			]),
		);
		expect(result.status).toBe('admitted');
		expect(result.metrics.points).toBe(2);
	});
	it('handles literal pen/brush lifetimes and stock selections', () => {
		const result = inspect(
			emf([
				record(38, [1, 0, 0, 0, 0]),
				record(39, [2, 0, 0xffffff, 0]),
				record(37, [1]),
				record(37, [2]),
				record(43, [0, 0, 10, 10]),
				record(37, [0x80000007]),
				record(37, [0x80000000]),
				record(40, [1]),
				record(40, [2]),
				record(48, [0x8000000f]),
			]),
		);
		expect(result.status).toBe('admitted');
		expect(result.metrics.objectsCreated).toBe(2);
	});
	it('accepts inert text state without permitting text or font records', () => {
		expect(
			inspect(
				emf([
					record(18, [1]),
					record(19, [2]),
					record(20, [13]),
					record(22, [24]),
					record(24, [0]),
					record(25, [0xffffff]),
				]),
			).status,
		).toBe('admitted');
		for (const type of [76, 82, 84, 96, 98, 114, 115])
			expect(codes(emf([record(type)]))).toContain('unsupported-record');
	});
	it('admits only the exact original-equivalent opaque WMF metadata convention', () => {
		const b = checksumComment(emf([sourceComment(), record(43, [0, 0, 10, 10])]));
		expect(inspect(b).status).toBe('admitted');
		expect(inspect(b).metrics.commentBytes).toBe(44);
		expect(codes(setWord(b, b.length - 8, 0))).toContain('embedded-wmf');
		const moved = checksumComment(emf([record(33), sourceComment(), record(34, [-1])]), 116);
		expect(codes(moved)).toContain('embedded-wmf');
	});
	it('rejects every EMF+ comment without descending into the stream', () => {
		const result = inspect(emf([record(70, [16, 0x2b464d45, 0xffffffff, 0xffffffff, 0xffffffff])]));
		expect(result.status).toBe('unsupported');
		expect(result.diagnostics[0]?.code).toBe('emf-plus');
		expect(result.metrics.records).toBe(3);
	});
	it.each([0, 0x43494447, 0x12345678])('rejects unresolved private/public comment %i', (id) => {
		expect(codes(emf([record(70, [8, id, 0x40000004])]))).toContain('comment-subtype');
	});
	it('rejects unresolved palette colors, styled pens, raster operations and MM_TEXT extents', () => {
		expect(codes(emf([record(38, [1, 6, 3, 0, 0x02938f6c])]))).toEqual(
			expect.arrayContaining(['pen-style', 'palette-color']),
		);
		expect(codes(emf([record(20, [7])]))).toContain('parameter');
		expect(codes(emf([record(9, [10, 10])]))).toContain('mapping-extents-fidelity');
	});
	it('counts clip ancestry and resets it with omitted RGN_COPY', () => {
		const result = inspect(
			emf([
				record(30, [0, 0, 90, 90]),
				record(33),
				record(30, [0, 0, 80, 80]),
				polygon16(triangle),
				record(34, [-1]),
				record(75, [0, 5]),
				polygon16(triangle),
			]),
		);
		expect(result.status).toBe('admitted');
		expect(result.metrics).toMatchObject({ clipOperations: 3, peakClipDepth: 2 });
	});
	it('does not silently skip nonempty or union regions', () => {
		const r = record(75, [48, 5, 32, 1, 1, 16, 0, 0, 10, 10, 0, 0, 10, 10]);
		expect(inspect(emf([r])).status).toBe('unsupported');
		expect(inspect(emf([r])).metrics.regionRects).toBe(1);
		expect(codes(emf([record(75, [0, 2])]))).toContain('region-data');
	});
	it('bounds diagnostic output even for repeated unknown records', () => {
		const result = inspect(emf(Array.from({ length: 200 }, () => record(500))), {
			maxDiagnostics: 3,
		});
		expect(result.status).toBe('unsupported');
		expect(result.diagnostics).toHaveLength(3);
		expect(result.omittedDiagnostics).toBe(197);
	});
	it('rechecks mutable inputs and reports no converter authorization', () => {
		const b = emf([record(43, [0, 0, 10, 10])]);
		expect(inspect(b).status).toBe('admitted');
		new DataView(b.buffer).setUint32(108, 76, true);
		expect(inspect(b).status).toBe('unsupported');
	});
	it('makes every exported default a hard ceiling', () => {
		expect(Object.isFrozen(VISIO_EMF_ADMISSION_LIMITS)).toBe(true);
		for (const key of Object.keys(
			VISIO_EMF_ADMISSION_LIMITS,
		) as (keyof typeof VISIO_EMF_ADMISSION_LIMITS)[]) {
			for (const value of [0, -1, NaN, Infinity, 0.5, VISIO_EMF_ADMISSION_LIMITS[key] + 1])
				expect(inspect(emf([]), { [key]: value }).status).toBe('invalid');
		}
	});
});
