import type { RegionMapOptions } from './types';
import { WORLD_REGIONS, type RegionDef } from './world-regions';
export interface RegionBounds {
	minX: number;
	minY: number;
	maxX: number;
	maxY: number;
}

export function regionBounds(region: RegionDef): RegionBounds {
	const coordinates = region.path.match(/-?\d+(?:\.\d+)?/gu)?.map(Number) ?? [];
	let minX = Infinity,
		minY = Infinity,
		maxX = -Infinity,
		maxY = -Infinity;
	for (let index = 0; index < coordinates.length; index += 2) {
		const x = coordinates[index] ?? 0,
			y = coordinates[index + 1] ?? 0;
		minX = Math.min(minX, x);
		maxX = Math.max(maxX, x);
		minY = Math.min(minY, y);
		maxY = Math.max(maxY, y);
	}
	return { minX, minY, maxX, maxY };
}

export function regionViewBounds(
	viewedRegionType: RegionMapOptions['viewedRegionType'],
	regionValues: ReadonlyMap<string, unknown>,
	regions: readonly RegionDef[] = WORLD_REGIONS,
): RegionBounds {
	if (!viewedRegionType || viewedRegionType === 'world' || regionValues.size === 0) {
		return { minX: 0, minY: 0, maxX: 1000, maxY: 500 };
	}
	const matched = regions.filter((region) => regionValues.has(region.code));
	const targets = viewedRegionType === 'countryRegion' ? matched.slice(0, 1) : matched;
	if (targets.length === 0) {
		return { minX: 0, minY: 0, maxX: 1000, maxY: 500 };
	}
	const bounds = targets.map(regionBounds);
	const padding = 10;
	return {
		minX: Math.max(0, Math.min(...bounds.map((item) => item.minX)) - padding),
		minY: Math.max(0, Math.min(...bounds.map((item) => item.minY)) - padding),
		maxX: Math.min(1000, Math.max(...bounds.map((item) => item.maxX)) + padding),
		maxY: Math.min(500, Math.max(...bounds.map((item) => item.maxY)) + padding),
	};
}

// ─────────────────────────────────────────────────────────────────────────────
// Path scaling
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Pre-scale and translate each coordinate pair in a simple SVG path `d`
 * attribute (M/L/Z commands only, space/comma delimited).
 *
 * This avoids needing a `transform` attribute on `SvgPath` which doesn't exist
 * in the existing primitive schema.
 */
export function scalePathD(d: string, scale: number, dx: number, dy: number): string {
	// Tokenise into command letters and numbers by MATCHING each token type
	// directly, rather than splitting on whitespace/comma delimiters: every
	// `WORLD_REGIONS` path is written in compact SVG syntax, where a command
	// letter directly abuts its first coordinate ("M130,160") and the closing
	// "Z" of one subpath directly abuts the last number before it ("130,195Z").
	// A delimiter-based split leaves "M130" and "195Z" as single un-split
	// tokens; `parseFloat("195Z")` silently truncates to 195 without ever
	// yielding a "Z", so the subpath is never closed and every coordinate
	// after it pairs up one token late, corrupting the rest of the shape and
	// eventually running the loop out of tokens mid-pair, i.e. exactly the
	// "attribute d: Unexpected end of attribute" the browser reports. Matching
	// letters and numbers as their own tokens (instead of splitting the
	// string on delimiters) handles both spaced and compact syntax alike.
	const tokens = d.match(/[A-Za-z]|-?\d+(?:\.\d+)?/gu) ?? [];
	const out: string[] = [];
	let i = 0;
	while (i < tokens.length) {
		const tok = tokens[i];
		if (tok === undefined) {
			i++;
			continue;
		}
		// Command letter (M, L, Z, etc.)
		if (/^[A-Za-z]$/u.test(tok)) {
			out.push(tok);
			i++;
			continue;
		}
		// Coordinate pair: tok = x value, tokens[i+1] = y value.
		const xRaw = parseFloat(tok);
		const yRaw = parseFloat(tokens[i + 1] ?? '');
		if (!Number.isNaN(xRaw) && !Number.isNaN(yRaw)) {
			out.push(`${(xRaw * scale + dx).toFixed(2)},${(yRaw * scale + dy).toFixed(2)}`);
			i += 2;
		} else {
			out.push(tok);
			i++;
		}
	}
	return out.join(' ');
}
