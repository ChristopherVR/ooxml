/**
 * OOXML DrawingML geometry path evaluation.
 *
 * Evaluates custom geometry definitions (a:custGeom) with formula-resolved
 * coordinates, producing SVG path data strings.
 */

import type { XmlObject } from '../types';
import { orderedPathCommandEntries } from './custom-geometry-command-order';
import { resolveCoordinate } from '../../../geometry/guide-formula-api';
import { ooxmlArcToSvg } from '../../../geometry/ooxml-arc';

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Evaluate a complete custom geometry definition (a:custGeom) with
 * formula-resolved coordinates, producing an SVG path data string.
 *
 * This handles the case where path coordinates reference guide names
 * instead of being plain numbers.
 *
 * @param pathNodes - Array of `a:path` XML objects from `a:pathLst`.
 * @param variables - Fully resolved variable context from evaluateGuides.
 * @param ensureArray - Helper to normalize XML nodes to arrays.
 * @returns SVG path data string and coordinate-space dimensions.
 */
export function evaluateGeometryPaths(
	pathNodes: ReadonlyArray<Record<string, unknown>>,
	variables: Map<string, number>,
	ensureArray: (val: unknown) => unknown[],
): { pathData: string; pathWidth: number; pathHeight: number } | null {
	let fullPathData = '';
	// Coordinate-space dimensions (from @_w / @_h on the first path that specifies them)
	const pathWidth =
		pathNodes
			.map((path) => Number.parseInt(String(path['@_w'] ?? '0'), 10))
			.find((width) => width > 0) ??
		variables.get('w') ??
		0;
	const pathHeight =
		pathNodes
			.map((path) => Number.parseInt(String(path['@_h'] ?? '0'), 10))
			.find((height) => height > 0) ??
		variables.get('h') ??
		0;

	for (const path of pathNodes) {
		// Each path element may declare its own coordinate-space dimensions
		const w = Number.parseInt(String(path['@_w'] ?? '0'), 10);
		const h = Number.parseInt(String(path['@_h'] ?? '0'), 10);
		const scaleX = pathWidth > 0 && w > 0 ? pathWidth / w : 1;
		const scaleY = pathHeight > 0 && h > 0 ? pathHeight / h : 1;
		const resolveX = (value: string | number | undefined) =>
			resolveCoordinate(value, variables) * scaleX;
		const resolveY = (value: string | number | undefined) =>
			resolveCoordinate(value, variables) * scaleY;

		const commands: string[] = [];
		// Track current pen position for arcTo conversion (arcTo needs the
		// current position to derive the implicit ellipse center)
		let penX = 0;
		let penY = 0;
		// Track the most recent moveTo position for close commands
		let moveX = 0;
		let moveY = 0;

		for (const [key, item] of orderedPathCommandEntries(path as XmlObject, ensureArray)) {
			if (!item || typeof item !== 'object') {
				if (key === 'a:close') {
					commands.push('Z');
					penX = moveX;
					penY = moveY;
				}
				continue;
			}

			const record = item as Record<string, unknown>;

			if (key === 'a:moveTo') {
				const pt = record['a:pt'] as Record<string, unknown> | undefined;
				if (pt) {
					const x = resolveX(pt['@_x'] as string | number | undefined);
					const y = resolveY(pt['@_y'] as string | number | undefined);
					commands.push(`M ${x} ${y}`);
					penX = x;
					penY = y;
					moveX = x;
					moveY = y;
				}
			} else if (key === 'a:lnTo') {
				const pt = record['a:pt'] as Record<string, unknown> | undefined;
				if (pt) {
					const x = resolveX(pt['@_x'] as string | number | undefined);
					const y = resolveY(pt['@_y'] as string | number | undefined);
					commands.push(`L ${x} ${y}`);
					penX = x;
					penY = y;
				}
			} else if (key === 'a:cubicBezTo') {
				const pts = ensureArray(record['a:pt']) as Array<Record<string, unknown>>;
				if (pts.length === 3) {
					const coords = pts.map((pt) => ({
						x: resolveX(pt['@_x'] as string | number | undefined),
						y: resolveY(pt['@_y'] as string | number | undefined),
					}));
					commands.push(
						`C ${coords[0].x} ${coords[0].y} ${coords[1].x} ${coords[1].y} ${coords[2].x} ${coords[2].y}`,
					);
					penX = coords[2].x;
					penY = coords[2].y;
				}
			} else if (key === 'a:quadBezTo') {
				const pts = ensureArray(record['a:pt']) as Array<Record<string, unknown>>;
				if (pts.length === 2) {
					const coords = pts.map((pt) => ({
						x: resolveX(pt['@_x'] as string | number | undefined),
						y: resolveY(pt['@_y'] as string | number | undefined),
					}));
					commands.push(`Q ${coords[0].x} ${coords[0].y} ${coords[1].x} ${coords[1].y}`);
					penX = coords[1].x;
					penY = coords[1].y;
				}
			} else if (key === 'a:arcTo') {
				const wR = resolveX(record['@_wR'] as string | number | undefined);
				const hR = resolveY(record['@_hR'] as string | number | undefined);
				const stAng = resolveCoordinate(
					record['@_stAng'] as string | number | undefined,
					variables,
				);
				const swAng = resolveCoordinate(
					record['@_swAng'] as string | number | undefined,
					variables,
				);

				const result = ooxmlArcToSvg(wR, hR, stAng, swAng, penX, penY);
				if (result) {
					commands.push(result.svg);
					penX = result.endX;
					penY = result.endY;
				}
			} else if (key === 'a:close') {
				commands.push('Z');
				penX = moveX;
				penY = moveY;
			}
		}

		if (commands.length > 0) {
			fullPathData += `${commands.join(' ')} `;
		}
	}

	const trimmed = fullPathData.trim();
	if (trimmed === '') {
		return null;
	}

	return {
		pathData: trimmed,
		pathWidth,
		pathHeight,
	};
}
