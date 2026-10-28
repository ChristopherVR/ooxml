import { describe, it, expect } from 'vitest';
import { Deflate, deflateSync, strToU8 } from 'fflate';
import { decodeGeoRing, readGeographyCache, inflateGeoCache } from './cache';
import { projectGeoPoint } from './geographic-projection';
const encoded = 'vx1vilihnM6hR7mEl2Q';
// Published quantized Bing PCA example, not redistributed provider boundaries.
const clear = `<cx:clear xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:geoData entityId="sample" entityName="Sample"><cx:geoPolygons><cx:geoPolygon polygonId="1" numPoints="4" pcaRings="1,${encoded}"/></cx:geoPolygons><cx:copyrights><cx:copyright>Test geometry</cx:copyright></cx:copyrights></cx:geoData></cx:clear>`;
describe('document geography cache', () => {
	it('decodes the published PCA vector with 51-bit values and signed deltas', () => {
		expect(decodeGeoRing(encoded)).toEqual([
			{ latitude: 35.89431, longitude: -110.72522 },
			{ latitude: 35.89393, longitude: -110.72578 },
			{ latitude: 35.89374, longitude: -110.72606 },
			{ latitude: 35.89337, longitude: -110.72662 },
		]);
	});
	it('reads clear and native raw-DEFLATE caches without inventing bbox polygons', () => {
		const bytes = deflateSync(strToU8(clear));
		const binary = btoa(String.fromCharCode(...bytes));
		const zipped = `<cx:geoCache provider="test"><cx:binary>${binary}</cx:binary></cx:geoCache>`;
		expect(readGeographyCache(zipped)).toEqual(readGeographyCache(clear));
		expect(readGeographyCache(clear).regions[0]?.copyright).toEqual(['Test geometry']);
		expect(
			readGeographyCache(clear.replace(/<cx:geoPolygons>[\s\S]*?<\/cx:geoPolygons>/u, '')).regions,
		).toEqual([]);
	});
	it('refuses corrupt rings and bounds decompression before parsing', () => {
		expect(decodeGeoRing('!')).toBeUndefined();
		expect(readGeographyCache(clear.replace(encoded, 'invalid!')).regions).toEqual([]);
		expect(readGeographyCache(clear.replace(encoded, 'invalid!')).issues.length).toBeGreaterThan(0);
		const bomb = btoa(String.fromCharCode(...deflateSync(strToU8('x'.repeat(1024)))));
		expect(() => inflateGeoCache(bomb, 64)).toThrow(/limit/u);
	});
	it('accepts Office sync-flushed streams without a final DEFLATE block', () => {
		const chunks: Uint8Array[] = [];
		const stream = new Deflate((chunk) => chunks.push(chunk));
		stream.push(strToU8(clear), false);
		stream.flush(true);
		const bytes = chunks.flatMap((chunk) => [...chunk]);
		expect(inflateGeoCache(btoa(String.fromCharCode(...bytes)))).toBe(clear);
	});
	it('shares the remaining point budget across every ring before allocating them', () => {
		const multi = clear
			.replace('numPoints="4"', 'numPoints="8"')
			.replace(`1,${encoded}`, `1,${encoded},${encoded}`);
		expect(readGeographyCache(multi, 8).regions[0]?.rings).toHaveLength(2);
		const bounded = readGeographyCache(multi, 6);
		expect(bounded.regions).toEqual([]);
		expect(bounded.issues).toContain('Invalid compressed geography polygon.');
	});
	it.each(['mercator', 'miller', 'robinson', 'albers'])(
		'projects cached coordinates with %s',
		(projection) => {
			expect(
				projectGeoPoint({ latitude: 30, longitude: 20 }, projection)?.every(Number.isFinite),
			).toBe(true);
			expect(projectGeoPoint({ latitude: 0, longitude: 0 }, 'unknown')).toBeUndefined();
		},
	);
});
