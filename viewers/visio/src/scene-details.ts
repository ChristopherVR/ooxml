import type { VisioShape } from 'ooxml-core/visio';

export interface DetailBudget {
	rows: number;
	characters: number;
}
/** Runtime model checks prevent implicit DOM stringification of host-provided objects. */
export function assertShapeDetails(shape: VisioShape, budget: DetailBudget): void {
	const data = shape.shapeData ?? [],
		links = shape.hyperlinks ?? [];
	if (
		!Array.isArray(data) ||
		!Array.isArray(links) ||
		data.length + links.length > 1024 ||
		(budget.rows += data.length + links.length) > 100_000
	)
		throw new Error('The scene exceeds shape metadata row limits.');
	const strings = (record: Record<string, unknown>, names: readonly string[], required = false) => {
		for (const name of names)
			if ((required || record[name] !== undefined) && typeof record[name] !== 'string')
				throw new Error('The scene has invalid shape metadata text.');
	};
	const booleans = (record: Record<string, unknown>, names: readonly string[]) => {
		for (const name of names)
			if (record[name] !== undefined && typeof record[name] !== 'boolean')
				throw new Error('The scene has invalid shape metadata flags.');
	};
	const count = (record: object, skipTarget = false) => {
		const values = Object.entries(record);
		if (values.length > 32) throw new Error('The scene has excessive shape metadata fields.');
		for (const [name, value] of values) {
			if (skipTarget && name === 'target') continue;
			if (typeof value === 'string') {
				if (value.length > 8192 || (budget.characters += value.length) > 5_000_000)
					throw new Error('The scene exceeds shape metadata string limits.');
			} else if (typeof value === 'number') {
				if (!Number.isFinite(value)) throw new Error('The scene has invalid numeric metadata.');
			} else if (value !== undefined && value !== null && typeof value !== 'boolean')
				throw new Error('The scene has invalid shape metadata values.');
		}
	};
	for (const row of data) {
		const record = row as unknown as Record<string, unknown>;
		strings(record, ['id', 'name', 'valueKind'], true);
		strings(record, [
			'label',
			'rawType',
			'rawValue',
			'unit',
			'error',
			'format',
			'prompt',
			'sortKey',
		]);
		booleans(record, ['invisible']);
		count(row);
	}
	for (const row of links) {
		const record = row as unknown as Record<string, unknown>;
		strings(record, ['id', 'name'], true);
		strings(record, ['description', 'address', 'subAddress', 'sortKey']);
		booleans(record, ['default', 'newWindow', 'invisible']);
		count(row, true);
		if (!row.target || typeof row.target !== 'object' || Array.isArray(row.target))
			throw new Error('The scene has invalid hyperlink metadata.');
		const target = row.target as unknown as Record<string, unknown>;
		const key =
			row.target.kind === 'external'
				? 'href'
				: row.target.kind === 'internal'
					? 'subAddress'
					: row.target.kind === 'unresolved'
						? 'reason'
						: undefined;
		if (!key) throw new Error('The scene has invalid hyperlink metadata.');
		strings(target, [key], true);
		count(row.target);
	}
}
