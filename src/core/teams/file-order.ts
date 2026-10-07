import type { FileEntry } from './view';

export type FileSortField = 'name' | 'ts' | 'author';
export type FileSortDirection = 'ascending' | 'descending';

/** Sort a copy, retaining source order for ties and every attachment's identity. */
export function sortFiles(
	files: readonly FileEntry[],
	field: FileSortField,
	direction: FileSortDirection,
): FileEntry[] {
	const sign = direction === 'ascending' ? 1 : -1;
	return [...files].sort((left, right) => {
		const order =
			field === 'ts'
				? left.ts - right.ts
				: left[field].localeCompare(right[field], undefined, {
						numeric: true,
						sensitivity: 'base',
					});
		return order * sign;
	});
}
