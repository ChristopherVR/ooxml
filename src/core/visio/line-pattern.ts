import type { VisioStyle } from './model';
import { number, type Cells, type Report } from './sheet';

/**
 * MS-VSDX 2.4.4.180 defines 0 transparent, 1 solid, pictured built-ins 2-23,
 * and 254 custom-master patterns. It does not specify numerical dash lengths:
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/f718c500-6066-4fcc-8b66-4bb5015f133f
 * These lengths match native Visio 16 SVG exports for every built-in pattern,
 * all three caps and 1/3-point strokes. See record-visio-line-patterns.ps1.
 * Native square-cap dots have a fixed 0.01-point length, stored separately
 * from the stroke-width ratios so thin and zero-width strokes remain finite.
 */
function builtinDash(pattern: number): number[] {
	if (pattern === 23) return [2, 1];
	const band = Math.floor((pattern - 2) / 7);
	const motif = (pattern - 2) % 7;
	const dash = [8, 4, 16][band]!;
	const gap = [4, 2, 8][band]!;
	const long = [20, 10, 40][band]!;
	const lengths = [
		[dash],
		[1],
		[dash, 1],
		[dash, 1, 1],
		[dash, dash, 1],
		[long, dash],
		[long, dash, dash],
	][motif]!;
	return lengths.flatMap((length) => [length, gap]);
}

/** Cached values win over formulas and themes; formulas are never evaluated. */
export function linePattern(
	cells: Cells,
	report: Report,
	cap: VisioStyle['lineCap'] = 'round',
): Pick<VisioStyle, 'linePattern' | 'lineDash' | 'lineDashDotLength'> {
	const raw = cells.get('LinePattern')?.value;
	if (raw === 'Themed') {
		report(
			'unresolved-line-pattern',
			'The requested theme line pattern is not resolved; a solid fallback was used.',
		);
		return { linePattern: 1 };
	}
	// Check before trimming so whitespace cannot evade the numeric-cache bound.
	if (raw !== undefined && raw.length > 128) {
		report('missing-cached-value', 'Cell LinePattern exceeds the bounded numeric cache length.');
		return { linePattern: 1 };
	}
	const pattern = number(cells, 'LinePattern', 1, report);
	if (Number.isInteger(pattern) && pattern >= 2 && pattern <= 23) {
		const lineDash = builtinDash(pattern).map((value, index) =>
			cap === 'butt' ? value : value + (index % 2 ? 1 : -1),
		);
		return {
			linePattern: pattern,
			lineDash,
			...(cap === 'square' && lineDash.includes(0) ? { lineDashDotLength: 0.01 / 72 } : {}),
		};
	}
	if (pattern === 254)
		report(
			'unsupported-custom-line-pattern',
			'Custom master-based line patterns are unsupported; a solid fallback is used.',
		);
	else if (pattern !== 0 && pattern !== 1) {
		report(
			'invalid-line-pattern',
			'The cached line-pattern value is outside the supported enumeration; a solid fallback was used.',
		);
		return { linePattern: 1 };
	}
	return { linePattern: pattern };
}
