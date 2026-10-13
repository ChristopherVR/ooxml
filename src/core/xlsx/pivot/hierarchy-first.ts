import type { CellValue, PivotTable } from '../model';
import type { PivotGrid } from './aggregate';
import { pivotItemLabel, pivotValueKey } from './data';
import { pivotMeasures, pivotMeasureNumber } from './measures';

type Entry = NonNullable<PivotGrid['rowEntries']>[number];
/** First Values reuse the hierarchy's indexed tuples and column aggregate buckets. */
export function appendFirstPivotValues(
	pivot: PivotTable,
	context: {
		rowItems: CellValue[];
		secondRowItems: CellValue[];
		columnItems: CellValue[];
		children: Map<string, Set<string>>;
		innerIndices: Map<string, number>;
		groups: Map<string, CellValue[][]>;
		details: Map<string, CellValue[][]>;
		buckets: Map<string, CellValue[][]>;
		width: number;
		rowLabels: number;
		top: boolean;
		bottom: boolean;
		append(labels: CellValue[], numbers: CellValue[], entry: Entry, indent?: number): void;
	},
): void {
	const {
		rowItems,
		secondRowItems,
		columnItems,
		children,
		innerIndices,
		groups,
		details,
		buckets,
		width,
		rowLabels,
		top,
		bottom,
		append,
	} = context;
	const compact = pivot.layout === 'compact';
	const tabular = !pivot.layout || pivot.layout === 'tabular';
	const blank = () => Array<CellValue>(width - rowLabels).fill(null);
	pivotMeasures(pivot).forEach((measure, m) => {
		const numbers = (keys: string[], records: CellValue[][]) => [
			...columnItems.map((column) =>
				pivotMeasureNumber(
					pivot,
					buckets.get(JSON.stringify([...keys, pivotValueKey(column)])) ?? [],
					measure,
				),
			),
			pivotMeasureNumber(pivot, records, measure),
		];
		if (!tabular)
			append([measure.caption, ...Array<CellValue>(rowLabels - 1).fill(null)], blank(), {
				indices: [m],
				measure: m,
			});
		let first = true;
		rowItems.forEach((outer, outerIndex) => {
			const key = pivotValueKey(outer);
			const group = groups.get(key) ?? [];
			if (!tabular)
				append(
					compact ? [pivotItemLabel(outer)] : [null, pivotItemLabel(outer), null],
					top ? numbers([key], group) : blank(),
					{ indices: [m, outerIndex], measure: m },
					compact ? 1 : 0,
				);
			const innerKeys = [...children.get(key)!].sort(
				(a, b) => innerIndices.get(a)! - innerIndices.get(b)!,
			);
			innerKeys.forEach((innerKey, child) => {
				const innerIndex = innerIndices.get(innerKey)!;
				append(
					compact
						? [pivotItemLabel(secondRowItems[innerIndex]!)]
						: [
								tabular && first ? measure.caption : null,
								pivot.repeatLabels || (tabular && child === 0) ? pivotItemLabel(outer) : null,
								pivotItemLabel(secondRowItems[innerIndex]!),
							],
					numbers([key, innerKey], details.get(JSON.stringify([key, innerKey]))!),
					{ indices: [m, outerIndex, innerIndex], measure: m },
					compact ? 2 : 0,
				);
				first = false;
			});
			if (bottom)
				append(
					compact
						? [`${String(pivotItemLabel(outer))} Total`]
						: [null, `${String(pivotItemLabel(outer))} Total`, null],
					numbers([key], group),
					{ indices: [m, outerIndex], measure: m, type: 'default' },
					compact ? 1 : 0,
				);
		});
	});
}
