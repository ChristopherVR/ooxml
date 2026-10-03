import { describe, expect, it, vi } from 'vitest';
import type { VsdDocument } from '@christophervr/ole2';
import { parseVsd } from '@christophervr/ole2';
import { parseLegacyVsd } from './legacy.js';
import { loadVisio } from './load.js';
import { parseVsdx } from './parser.js';
import { fixture, shape, rectangle } from './test-fixtures.js';

vi.mock('@christophervr/ole2', async (importOriginal) => ({
	...(await importOriginal<typeof import('@christophervr/ole2')>()),
	parseVsd: vi.fn(),
}));

function drawing(
	overrides: Record<string, unknown> = {},
	page: Record<string, unknown> = {},
): VsdDocument {
	return {
		pages: [
			{
				id: 0,
				width: 8,
				height: 11,
				scale: 1,
				background: false,
				shapes: [
					{
						id: 7,
						kind: 'shape',
						coordinateSpace: 'shape-local',
						transformCoordinateSpace: 'parent-local',
						parentId: 0xffffffff,
						masterPageId: 0xffffffff,
						masterShapeId: 0xffffffff,
						unsupportedGeometry: false,
						text: 'Hello <script> & world\n',
						transform: {
							pinX: 2,
							pinY: 3,
							width: 4,
							height: 2,
							localPinX: 2,
							localPinY: 1,
							angle: 0,
							flipX: false,
							flipY: false,
						},
						geometry: [
							{ kind: 'moveTo', id: 0, x: 0, y: 0 },
							{ kind: 'lineTo', id: 1, x: 4, y: 2 },
						],
						...overrides,
					},
				],
				...page,
			},
		],
	} as unknown as VsdDocument;
}
const cfbHeader = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
const normalizeLegacyVsd = (
	document: VsdDocument,
	options: Parameters<typeof parseLegacyVsd>[1] = {},
) => {
	vi.mocked(parseVsd).mockReturnValue(document);
	return parseLegacyVsd(cfbHeader, options);
};

describe('legacy VSD scene adapter', () => {
	it('maps explicit inch geometry and inert text with honest style diagnostics', () => {
		const document = normalizeLegacyVsd(drawing()),
			page = document.pages[0]!,
			shape = page.shapes[0]!;
		expect(document.format).toBe('vsd');
		expect(page).toMatchObject({
			id: '0',
			width: 8,
			height: 11,
			isBackground: false,
			connectors: [],
		});
		expect(shape.transform).toEqual([1, 0, -0, 1, 0, 2]);
		expect(shape.geometry).toEqual([{ path: 'M 0 0 L 4 2', fill: false, stroke: true }]);
		expect(shape.text.plainText).toBe('Hello <script> & world\n');
		expect(shape.text.runs[0]?.text).toBe(shape.text.plainText);
		expect(shape.style).toMatchObject({ fill: 'none', lineColor: '#000000' });
		expect(document.diagnostics.map((d) => d.code)).toEqual([
			'legacy-vsd-stored-values',
			'legacy-vsd-style-fallback',
			'legacy-vsd-text-fallback',
		]);
	});
	it('uses the existing shape transform implementation for rotation, flips and explicit local pins', () => {
		const transform = {
			pinX: 5,
			pinY: 6,
			width: 4,
			height: 2,
			localPinX: 1,
			localPinY: 2,
			angle: Math.PI / 2,
			flipX: true,
			flipY: false,
		};
		const matrix = normalizeLegacyVsd(drawing({ transform })).pages[0]!.shapes[0]!.transform;
		expect(matrix[0]).toBeCloseTo(0);
		expect(matrix[1]).toBeCloseTo(-1);
		expect(matrix[2]).toBeCloseTo(-1);
		expect(matrix[3]).toBeCloseTo(0);
		expect(matrix[4]).toBeCloseTo(7);
		expect(matrix[5]).toBeCloseTo(7);
	});
	it.each([
		{ kind: 'group' },
		{ kind: 'foreign' },
		{ parentId: 3 },
		{ masterPageId: 1 },
		{ masterShapeId: 7 },
		{ parentId: undefined },
		{ masterPageId: undefined },
		{ masterShapeId: undefined },
		{ unsupportedGeometry: true },
		{ geometry: [] },
		{ transform: undefined },
		{ coordinateSpace: 'page-local' },
		{ transformCoordinateSpace: 'page-local' },
		{ geometry: [{ kind: 'lineTo', x: 1, y: 1 }] },
		{ geometry: [{ kind: 'moveTo', x: NaN, y: 0 }] },
		{ text: 'unresolved\u001efield' },
		{ text: 'unresolved\ufffcobject' },
	])('refuses unsafe or unresolved shape geometry/content %j', (overrides) => {
		expect(() => normalizeLegacyVsd(drawing(overrides))).toThrow();
	});
	it.each([{ width: undefined }, { height: -1 }, { width: Infinity }])(
		'refuses missing/unsafe page dimensions %j',
		(page) => {
			expect(() => normalizeLegacyVsd(drawing({}, page))).toThrow('dimensions');
		},
	);
	it.each([undefined, 0, 2])('refuses unresolved/scaled page coordinates %s', (scale) => {
		expect(() => normalizeLegacyVsd(drawing({}, { scale }))).toThrow('drawing scale');
	});
	it('bounds geometry/text/diagnostics and rejects invalid limits', () => {
		expect(() => normalizeLegacyVsd(drawing(), { maxGeometryCommands: 1 })).toThrow(
			'geometry limit',
		);
		expect(() => normalizeLegacyVsd(drawing(), { maxTextCharacters: 1 })).toThrow('text limit');
		expect(() => normalizeLegacyVsd(drawing(), { maxShapes: -1 })).toThrow(
			'positive safe integers',
		);
		const diagnostics = normalizeLegacyVsd(drawing(), { maxDiagnostics: 1 }).diagnostics;
		expect(diagnostics).toHaveLength(1);
		expect(diagnostics[0]?.code).toBe('diagnostics-truncated');
	});
	it('refuses arithmetic overflow and enforces the decoder runtime check', () => {
		const transform = {
			pinX: -Number.MAX_VALUE,
			pinY: 0,
			width: 1,
			height: 1,
			localPinX: Number.MAX_VALUE,
			localPinY: 0,
			angle: 0,
			flipX: false,
			flipY: false,
		};
		expect(() => normalizeLegacyVsd(drawing({ transform }))).toThrow('overflows');
		vi.mocked(parseVsd).mockReturnValue(drawing());
		const clock = vi.spyOn(Date, 'now').mockReturnValueOnce(0).mockReturnValueOnce(20_000);
		try {
			expect(() => parseLegacyVsd(cfbHeader)).toThrow('deadline');
		} finally {
			clock.mockRestore();
		}
	});
	it('delegates binary decoding to ole2 and applies input/runtime bounds', () => {
		vi.mocked(parseVsd).mockReturnValue(drawing());
		expect(parseLegacyVsd(cfbHeader).format).toBe('vsd');
		expect(parseVsd).toHaveBeenCalledWith(cfbHeader);
		expect(() => parseLegacyVsd(cfbHeader, { limits: { maxInputBytes: 1 } })).toThrow(
			'input exceeds',
		);
		vi.mocked(parseVsd).mockImplementation(() => {
			throw new Error('unsupported-version');
		});
		expect(() => parseLegacyVsd(cfbHeader)).toThrow('unsupported-version');
	});
	it('loads CFB as legacy VSD, retains ZIP VSDX behavior and parser options', async () => {
		vi.mocked(parseVsd).mockReturnValue(drawing());
		expect((await loadVisio(cfbHeader.buffer)).format).toBe('vsd');
		const bytes = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${shape('1', rectangle)}</Shapes>` }],
		});
		expect(await loadVisio(bytes)).toEqual(await parseVsdx(bytes));
		await expect(loadVisio(bytes, { maxGeometryCommands: 1 })).rejects.toThrow();
		await expect(loadVisio(Uint8Array.of(1, 2, 3))).rejects.toThrow('Expected a VSD');
	});
});
