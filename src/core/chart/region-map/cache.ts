// Reads CT_GeoCache clear/raw-DEFLATE XML; geometry belongs to the document's provider.
import { Inflate } from 'fflate';
import { parseXml, type XmlElement } from '../../xml';
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const LIMIT = 16 * 1024 * 1024;
const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-';
export interface GeoPoint {
	latitude: number;
	longitude: number;
}
export interface CachedGeoRegion {
	entityId: string;
	name: string;
	rings: GeoPoint[][];
	copyright: string[];
}
export interface GeographyCache {
	regions: CachedGeoRegion[];
	issues: string[];
}
/** Inverse of Microsoft's Bing point compression algorithm (51-bit Cantor pairs, 1e-5 degrees). */
export function decodeGeoRing(encoded: string, limit = 1_000_000): GeoPoint[] | undefined {
	const out: GeoPoint[] = [];
	let lat = 0,
		lon = 0,
		number = 0,
		shift = 0;
	for (const ch of encoded) {
		const value = ALPHABET.indexOf(ch);
		if (value < 0 || shift > 50) return undefined;
		number += (value & 31) * 2 ** shift;
		if (!Number.isSafeInteger(number)) return undefined;
		if (value >= 32) {
			shift += 5;
			continue;
		}
		const diagonal = Math.floor((Math.sqrt(8 * number + 1) - 1) / 2);
		const a = number - (diagonal * (diagonal + 1)) / 2,
			b = diagonal - a;
		lat += a % 2 ? -(a + 1) / 2 : a / 2;
		lon += b % 2 ? -(b + 1) / 2 : b / 2;
		if (Math.abs(lat) > 90e5 || Math.abs(lon) > 540e5 || out.length >= limit) return undefined;
		out.push({ latitude: lat / 1e5, longitude: lon / 1e5 });
		number = 0;
		shift = 0;
	}
	return shift || out.length < 3 ? undefined : out;
}
function elements(root: XmlElement, name: string): XmlElement[] {
	return Array.from(root.getElementsByTagNameNS(CX, name));
}
/** Bounded streaming inflate: never feed the whole compressed payload to an unbounded decoder. */
export function inflateGeoCache(binary: string, maxOutputBytes = LIMIT): string {
	if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes < 1 || maxOutputBytes > LIMIT)
		throw new Error('Invalid geography cache output limit.');
	if (binary.length > 12 * 1024 * 1024) throw new Error('Geography cache input limit exceeded.');
	const bytes = Uint8Array.from(atob(binary.replace(/\s/gu, '')), (c) => c.charCodeAt(0));
	const chunks: Uint8Array[] = [];
	let size = 0;
	const stream = new Inflate((chunk) => {
		size += chunk.length;
		if (size > maxOutputBytes) throw new Error('Geography cache output limit exceeded.');
		chunks.push(chunk);
	});
	for (let i = 0; i < bytes.length; i += 256)
		// Office writes a sync-flushed raw stream without BFINAL; XML completeness is validated below.
		stream.push(bytes.subarray(i, i + 256), false);
	const out = new Uint8Array(size);
	let offset = 0;
	for (const chunk of chunks) {
		out.set(chunk, offset);
		offset += chunk.length;
	}
	return new TextDecoder('utf-8', { fatal: true }).decode(out);
}
export function readGeographyCache(
	xml: string | undefined,
	pointLimit = 1_000_000,
): GeographyCache {
	const result: GeographyCache = { regions: [], issues: [] };
	if (!xml) return result;
	try {
		if (!Number.isSafeInteger(pointLimit) || pointLimit < 1 || pointLimit > 1_000_000)
			throw new Error('Invalid geography point limit.');
		if (xml.length > LIMIT) throw new Error('Geography cache input limit exceeded.');
		let root = parseXml(`<cx:root xmlns:cx="${CX}">${xml}</cx:root>`).documentElement;
		const binary = elements(root, 'binary')[0];
		if (binary) root = parseXml(inflateGeoCache(binary.textContent ?? '')).documentElement;
		let points = 0;
		for (const data of elements(root, 'geoData')) {
			const rings: GeoPoint[][] = [];
			for (const polygon of elements(data, 'geoPolygon')) {
				const parts = (polygon.getAttribute('pcaRings') ?? '').split(',');
				const version = Number(parts.shift());
				const expected = Number(polygon.getAttribute('numPoints'));
				if (version !== 1 || parts.length < 1 || !Number.isInteger(expected) || expected < 3) {
					result.issues.push('Invalid geography ring count.');
					continue;
				}
				const decoded: GeoPoint[][] = [];
				let polygonPoints = 0;
				for (const part of parts) {
					const ring = decodeGeoRing(
						part,
						Math.min(expected - polygonPoints, pointLimit - points - polygonPoints),
					);
					if (!ring) break;
					decoded.push(ring);
					polygonPoints += ring.length;
				}
				if (decoded.length !== parts.length) {
					result.issues.push('Invalid compressed geography polygon.');
					continue;
				}
				for (const ring of decoded) {
					rings.push(ring);
					points += ring.length;
				}
			}
			if (rings.length)
				result.regions.push({
					entityId: data.getAttribute('entityId') ?? '',
					name: data.getAttribute('entityName') ?? '',
					rings,
					copyright: elements(data, 'copyright').map((c) => c.textContent ?? ''),
				});
		}
	} catch (error) {
		result.issues.push(error instanceof Error ? error.message : String(error));
	}
	return result;
}
