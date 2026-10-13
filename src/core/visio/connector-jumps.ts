import type { VisioShape } from './model';
import type { Cells, Report } from './sheet';

/** Connector jump enums differ from PageSheet style enums: zero inherits the page. */
export function readConnectorJumps(
	cells: Cells,
	report: Report,
): Pick<
	VisioShape,
	'lineJumpCode' | 'lineJumpStyle' | 'lineJumpDirectionX' | 'lineJumpDirectionY'
> {
	const result: ReturnType<typeof readConnectorJumps> = {};
	for (const [key, name, maximum] of [
		['lineJumpCode', 'ConLineJumpCode', 4],
		['lineJumpStyle', 'ConLineJumpStyle', 9],
		['lineJumpDirectionX', 'ConLineJumpDirX', 2],
		['lineJumpDirectionY', 'ConLineJumpDirY', 2],
	] as const) {
		const cell = cells.get(name);
		if (!cell || cell.formula === 'Inh') continue;
		const raw = cell.value?.trim();
		const value =
			raw && raw.length <= 64 && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(raw)
				? Number(raw)
				: NaN;
		if (cell.error !== undefined || !Number.isInteger(value) || value < 0 || value > maximum) {
			report(
				'unsupported-connector-jump',
				`${name} has no supported cached value; the page setting is used.`,
			);
			continue;
		}
		result[key] = value;
	}
	return result;
}
