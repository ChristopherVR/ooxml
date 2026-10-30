import { parseXml } from '../../xml/index';
import { formatRange } from '../address';
import type { SortState, Worksheet } from '../model';
import { readSortState } from '../read/sort-state';
import { sameModel } from './snapshot';
import type { StyleWriter } from './styles';
import { el } from './xml-out';

/** `<sortState>` for a saved sort; colour levels write their format through `styles`. */
export function sortStateXml(state: SortState | undefined, styles?: StyleWriter): string {
	if (!state) return '';
	const levels = state.conditions
		.map((c) => {
			const dxfId = c.style && styles ? styles.dxfId(c.style) : c.dxfId;
			return el('sortCondition', {
				descending: c.descending || undefined,
				sortBy: c.sortBy && c.sortBy !== 'value' ? c.sortBy : undefined,
				ref: formatRange(c.ref),
				customList: c.customList,
				dxfId,
				iconSet: c.iconSet,
				iconId: c.iconId,
			});
		})
		.join('');
	return el(
		'sortState',
		{
			columnSort: state.columnSort || undefined,
			caseSensitive: state.caseSensitive || undefined,
			ref: formatRange(state.range),
		},
		levels,
	);
}

const bare = (state: SortState | undefined) =>
	state && {
		...state,
		conditions: state.conditions.map(({ style: _style, ...rest }) => rest),
	};

/**
 * The worksheet-level `<sortState>`. Empty while the model still equals the kept source element
 * (the save then slots the original back in unchanged).
 */
export function sheetSortStateXml(sheet: Worksheet, styles?: StyleWriter): string {
	if (!sheet.sortState) return '';
	const kept = sheet.preserved.get('sortState')?.[0];
	if (kept) {
		try {
			const source = readSortState(parseXml(kept, { label: 'XLSX worksheet' }).documentElement);
			if (sameModel(bare(source), bare(sheet.sortState))) return '';
		} catch {
			// Fall through and write the model.
		}
	}
	return sortStateXml(sheet.sortState, styles);
}
