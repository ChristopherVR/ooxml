import { readGeographyCache } from './cache';
import { projectGeoPoint } from './geographic-projection';
import { resolveRegionCode } from './alias';
import { WORLD_REGIONS, type RegionDef } from './world-regions';
import type { RegionMapOptions } from './types';
/** Document-provided boundaries take precedence; the bundled country silhouettes are explicitly approximate. */
export function regionMapGeometry(options: RegionMapOptions | undefined): {
	regions: RegionDef[];
	approximate: boolean;
	issues: string[];
	copyright: string[];
} {
	const cache = readGeographyCache(options?.geographyCacheXml);
	if (!cache.regions.length)
		return { regions: WORLD_REGIONS, approximate: true, issues: cache.issues, copyright: [] };
	// Native Office's automatic world view uses the rounded Robinson projection.
	const projection = options?.projectionType ?? 'robinson';
	const projected = cache.regions.map((region) => ({
		region,
		rings: region.rings.map((ring) => ring.map((p) => projectGeoPoint(p, projection))),
	}));
	if (projected.some((r) => r.rings.some((ring) => ring.some((p) => !p))))
		return {
			regions: [],
			approximate: false,
			issues: [...cache.issues, 'Unsupported geographic projection.'],
			copyright: [],
		};
	const all = projected.flatMap((r) => r.rings.flatMap((ring) => ring as [number, number][]));
	let minX = Infinity,
		maxX = -Infinity,
		minY = Infinity,
		maxY = -Infinity;
	for (const [x, y] of all) {
		minX = Math.min(minX, x);
		maxX = Math.max(maxX, x);
		minY = Math.min(minY, y);
		maxY = Math.max(maxY, y);
	}
	const factor = Math.min(1000 / Math.max(maxX - minX, 1e-12), 500 / Math.max(maxY - minY, 1e-12));
	const sx = factor,
		sy = factor,
		ox = (1000 - (maxX - minX) * factor) / 2,
		oy = (500 - (maxY - minY) * factor) / 2;
	const regions = projected.map(({ region, rings }): RegionDef => {
		const paths = rings.map((ring) =>
			ring.map((p) => {
				const [x, y] = p!;
				return [(x - minX) * sx + ox, (y - minY) * sy + oy];
			}),
		);
		const coords = paths.flat();
		let x = 0,
			y = 0;
		for (const p of coords) {
			x += p[0]!;
			y += p[1]!;
		}
		return {
			code: resolveRegionCode(region.name) ?? region.entityId,
			entityId: region.entityId,
			name: region.name,
			labelXY: [x / coords.length, y / coords.length],
			path: paths
				.map(
					(r) =>
						r.map((p, i) => `${i ? 'L' : 'M'}${p[0]!.toFixed(3)},${p[1]!.toFixed(3)}`).join(' ') +
						'Z',
				)
				.join(' '),
		};
	});
	return {
		regions,
		approximate: false,
		issues: cache.issues,
		copyright: [...new Set(cache.regions.flatMap((r) => r.copyright))],
	};
}
