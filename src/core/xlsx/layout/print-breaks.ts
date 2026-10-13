import { MAX_COL, MAX_ROW } from '../address';
import type { Worksheet } from '../model';
import { NS, children, parseXml } from '../../xml/index';

export type PrintAxis = 'row' | 'col';
export interface PrintBreak {
	before: number;
	min: number;
	max: number;
	manual: boolean;
}

/** Break IDs are zero-based positions before the next row or column. */
export function sheetPageBreaks(sheet: Worksheet, axis: PrintAxis): PrintBreak[] {
	const key = axis === 'row' ? 'rowBreaks' : 'colBreaks';
	const limit = axis === 'row' ? MAX_ROW : MAX_COL;
	const across = axis === 'row' ? MAX_COL : MAX_ROW;
	return (sheet.preserved.get(key) ?? []).flatMap((xml) =>
		children(parseXml(xml).documentElement, 'brk', NS.x).flatMap((node) => {
			const before = Number(node.getAttribute('id'));
			const min = Number(node.getAttribute('min') ?? 0);
			const max = Number(node.getAttribute('max') ?? 0);
			if (
				![before, min, max].every(Number.isInteger) ||
				before <= 0 ||
				before > limit ||
				min < 0 ||
				max < min ||
				max > across
			)
				return [];
			return [{ before, min, max, manual: /^(1|true)$/.test(node.getAttribute('man') ?? '') }];
		}),
	);
}
