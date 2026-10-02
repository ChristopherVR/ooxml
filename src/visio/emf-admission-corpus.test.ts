import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { inspectVisioEmfAdmission } from './emf-admission.js';

// Optional read-only external validation; never copies or redistributes corpus fixture bytes.
// Provenance: Apache POI test-data/diagram/60973.vsdx. Set VISIO_EMF_CORPUS_FILE explicitly.
const file = process.env.VISIO_EMF_CORPUS_FILE;
const inventory2 = {
	1: 1,
	9: 2,
	10: 1,
	11: 1,
	14: 1,
	17: 1,
	18: 1,
	20: 2,
	22: 2,
	24: 2,
	25: 2,
	30: 1,
	33: 3,
	34: 3,
	37: 32,
	38: 10,
	39: 2,
	40: 12,
	42: 9,
	45: 8,
	47: 8,
	48: 1,
	70: 1,
	75: 1,
};
const inventory3 = {
	1: 1,
	9: 2,
	10: 1,
	11: 1,
	14: 1,
	17: 1,
	18: 1,
	20: 2,
	22: 2,
	24: 2,
	25: 2,
	27: 2,
	30: 3,
	33: 5,
	34: 5,
	37: 24,
	38: 5,
	39: 5,
	40: 10,
	42: 2,
	48: 1,
	54: 2,
	70: 1,
	75: 1,
	86: 18,
};
const fixtures = [
	{
		name: 'image2.emf',
		bytes: 3736,
		sha256: '87543078d0617e08a4649d6141fc443418903f1f951417490d0348c91c6e49b4',
		inventory: inventory2,
		header: { pixelWidth: 83, pixelHeight: 50, declaredRecords: 107, declaredHandles: 6 },
		metrics: {
			objectsCreated: 12,
			peakLiveObjects: 5,
			clipOperations: 2,
			peakClipDepth: 1,
			peakStateDepth: 3,
			commentBytes: 1588,
		},
	},
	{
		name: 'image3.emf',
		bytes: 4216,
		sha256: '8093b61d382b703f07e499f239377b6d0092230d69d3f779a209cec682a89ee0',
		inventory: inventory3,
		header: { pixelWidth: 65, pixelHeight: 38, declaredRecords: 100, declaredHandles: 7 },
		metrics: {
			objectsCreated: 10,
			peakLiveObjects: 6,
			clipOperations: 4,
			peakClipDepth: 2,
			peakStateDepth: 4,
			commentBytes: 1856,
		},
	},
] as const;

describe.skipIf(!file)('explicit external Apache POI 60973.vsdx classic media', () => {
	it.each(fixtures)('characterizes $name without claiming render fidelity', async (expected) => {
		const zip = await JSZip.loadAsync(await readFile(file!));
		const bytes = await zip.file(`visio/media/${expected.name}`)!.async('uint8array');
		expect(bytes.byteLength).toBe(expected.bytes);
		expect(createHash('sha256').update(bytes).digest('hex')).toBe(expected.sha256);
		const result = inspectVisioEmfAdmission(bytes);
		expect(result).toMatchObject({
			status: 'unsupported',
			renderingEnabled: false,
			scanComplete: true,
			header: expected.header,
			metrics: expected.metrics,
		});
		expect(Object.fromEntries(result.recordTypes.map(({ type, count }) => [type, count]))).toEqual(
			expected.inventory,
		);
		expect(new Set(result.diagnostics.map((d) => d.code))).toEqual(
			new Set(['pen-style', 'palette-color']),
		);
		expect(result.diagnostics.every((d) => d.kind === 'unsupported')).toBe(true);
		expect(result.omittedDiagnostics).toBe(0);
	});
});
