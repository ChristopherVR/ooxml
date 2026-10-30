// Top 10 AutoFilter: show the top or bottom N items or percent of a column.
import { topFilterCriteria } from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, numberInput, row, select } from './fields';
import { showDialog } from './frame';

export interface Top10Props {
	col?: number;
}

export function openTop10Filter(
	ctx: EditorContext,
	props: Top10Props,
): Promise<boolean | undefined> {
	const t = target(ctx);
	const filter = t?.ws.autoFilter;
	if (!t || !filter) return Promise.resolve(undefined);
	const col = props.col ?? t.active.col;
	const current = filter.columns?.find((c) => c.offset === col - filter.range.start.col)?.top10;
	const side = select(
		ctx,
		[
			['top', 'Top'],
			['bottom', 'Bottom'],
		],
		current && !current.top ? 'bottom' : 'top',
	);
	side.setAttribute('aria-label', ctx.t('Top or bottom'));
	const count = numberInput(ctx, current?.value ?? 10, 1, 500);
	count.setAttribute('aria-label', ctx.t('Number of items'));
	const unit = select(
		ctx,
		[
			['items', 'Items'],
			['percent', 'Percent'],
		],
		current?.percent ? 'percent' : 'items',
	);
	unit.setAttribute('aria-label', ctx.t('Items or percent'));
	return showDialog<boolean>(ctx, {
		name: 'top10-filter',
		heading: 'Top 10 AutoFilter',
		body: [field(ctx, 'Show', row(ctx, side, count, unit))],
		opened: () => count.focus(),
		submit: () => {
			const value = Number(count.value);
			if (!Number.isFinite(value) || value < 1 || value > 500) {
				ctx.toast(ctx.t('Enter a number from 1 to 500.'), 'warning');
				return undefined;
			}
			t.session.setFilterCriteria(
				t.sheet,
				col,
				topFilterCriteria(side.value === 'top', unit.value === 'percent', Math.floor(value)),
			);
			return true;
		},
	});
}
