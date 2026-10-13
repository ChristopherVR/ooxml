import type { Cell } from './sheet';
type SourceCells = ReadonlyMap<string, Readonly<Cell>>;

// Parser row maps are immutable after inheritance resolution. Repeated markers share one
// bounded proof instead of repeatedly serializing an unusually large Field row.
const proofs = new WeakMap<SourceCells, string | undefined>();

/** Source-backed field cache identity, shared by the parser and atomic field writer. */
export function textFieldSourceKey(cells: SourceCells): string | undefined {
	if (proofs.has(cells)) return proofs.get(cells);
	if (cells.size > 64) {
		proofs.set(cells, undefined);
		return undefined;
	}
	const key = JSON.stringify(
		[...cells]
			.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
			.map(([name, cell]) => [
				name,
				cell.value ?? null,
				cell.formula ?? null,
				cell.unit ?? null,
				cell.error ?? null,
			]),
	);
	const bounded = key.length <= 2048 ? key : undefined;
	proofs.set(cells, bounded);
	return bounded;
}
