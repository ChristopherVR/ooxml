import type { VisioStyle } from './model.js';
import { number, type Cells, type Report } from './sheet.js';

/**
 * MS-VSDX 2.4.4.180 defines 0 transparent, 1 solid, pictured built-ins 2-23,
 * and 254 custom-master patterns. It does not specify numerical dash lengths:
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/f718c500-6066-4fcc-8b66-4bb5015f133f
 * The compatibility ratios below are inferred, not native-renderer measurements.
 * Corroborating implementation: LibreOffice/libvisio commit
 * 49fb9d3a9d21d4374cad782925e48c577a41f5be, src/lib/VSDContentCollector.cpp:3037-3196.
 * This is an original bounded normalization, not copied implementation code.
 */
function builtinDash(pattern: number): number[] {
	if (pattern === 23) return [2, 2];
	const band = Math.floor((pattern - 2) / 7);
	const motif = (pattern - 2) % 7;
	const dash = [6, 3, 11][band]!;
	const gap = [3, 2, 5][band]!;
	const long = [14, 7, 27][band]!;
	const longGap = [2, 2, 5][band]!;
	const lengths = [
		[dash],
		[1],
		[dash, 1],
		[dash, 1, 1],
		[dash, dash, 1],
		[long, dash],
		[long, dash, dash],
	][motif]!;
	return lengths.flatMap((length) => [length, motif >= 5 ? longGap : gap]);
}

/** Cached values win over formulas and themes; formulas are never evaluated. */
export function linePattern(
	cells: Cells,
	report: Report,
): Pick<VisioStyle, 'linePattern' | 'lineDash'> {
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
		report(
			'inferred-line-pattern',
			'Built-in line-pattern spacing uses inferred stroke-width ratios, not verified native spacing.',
		);
		return { linePattern: pattern, lineDash: builtinDash(pattern) };
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
