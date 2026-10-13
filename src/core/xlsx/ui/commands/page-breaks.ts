import type { Command } from '../commands';
import { icon } from './icons';
import { editing, target, viewing } from './util';

export function pageBreakCommands(): Command[] {
	return [
		viewing({
			id: 'page.preview',
			label: 'Print Preview',
			icon: icon('pageSetup'),
			run: (ctx) => void ctx.dialogs.open('print-preview'),
		}),
		...(['row', 'col'] as const).flatMap((axis) =>
			([true, false] as const).map((present) =>
				editing({
					id: `page.break-${axis}-${present ? 'insert' : 'remove'}`,
					label: present
						? axis === 'row'
							? 'Insert Row Page Break'
							: 'Insert Column Page Break'
						: axis === 'row'
							? 'Remove Row Page Break'
							: 'Remove Column Page Break',
					icon: icon('printArea'),
					lock: false,
					enabled: (ctx) => (target(ctx)?.active[axis] ?? 0) > 0,
					run: (ctx) => {
						const t = target(ctx);
						if (t) t.session.setPageBreak(t.sheet, axis, t.active[axis], present);
					},
				}),
			),
		),
		editing({
			id: 'page.breaks-reset',
			label: 'Reset All Page Breaks',
			icon: icon('printArea'),
			lock: false,
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.resetPageBreaks(t.sheet);
			},
		}),
	];
}
