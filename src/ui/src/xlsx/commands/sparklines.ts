import { selectedSparklineGroups } from 'ooxml-core/xlsx';
import { editing, icon, target, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';

function hasSparklines(ctx: EditorContext): boolean {
	const t = target(ctx);
	return !!t && selectedSparklineGroups(t.ws, t.ranges).length > 0;
}

export function sparklineCommands(): Command[] {
	return [
		...(
			[
				['line', 'Line Sparkline', 'chartLine'],
				['column', 'Column Sparkline', 'chartColumn'],
				['stacked', 'Win/Loss Sparkline', 'chartColumn'],
			] as const
		).map(([type, label, glyph]) =>
			editing({
				id: `insert.sparkline-${type}`,
				label,
				icon: icon(glyph),
				lock: 'objects',
				run: (ctx) => void ctx.dialogs.open('create-sparklines', { type }),
			}),
		),
		editing({
			id: 'sparkline.settings',
			label: 'Sparkline Settings',
			icon: icon('chartLine'),
			lock: 'objects',
			enabled: hasSparklines,
			run: (ctx) => void ctx.dialogs.open('sparkline-settings'),
		}),
		...(
			[
				['group', 'Group Sparklines', 'groupSparklines'],
				['ungroup', 'Ungroup Sparklines', 'ungroupSparklines'],
			] as const
		).map(([id, label, method]) =>
			editing({
				id: `sparkline.${id}`,
				label,
				icon: icon('chartLine'),
				lock: 'objects',
				enabled: hasSparklines,
				run: (ctx) => {
					const t = target(ctx);
					if (t) t.session[method](t.sheet, t.ranges);
				},
			}),
		),
		...(
			[
				['selected', false, 'Clear Selected Sparklines'],
				['groups', true, 'Clear Sparkline Groups'],
			] as const
		).map(([id, whole, label]) =>
			editing({
				id: `sparkline.clear-${id}`,
				label,
				icon: icon('clear'),
				lock: 'objects',
				enabled: hasSparklines,
				run: (ctx) => {
					const t = target(ctx);
					if (t) t.session.deleteSparklines(t.sheet, t.ranges, whole);
				},
			}),
		),
	];
}
