import type { Command } from '../commands';
import { icon } from './icons';
import { editing, target } from './util';

const DIMENSIONS = [
	['width', 'Width', 'fitToWidth'],
	['height', 'Height', 'fitToHeight'],
] as const;

/** A zero page count is Excel's Automatic setting. */
export function printScalingCommands(): Command[] {
	return [
		...DIMENSIONS.map(([axis, label, field]) =>
			editing({
				id: `page.fit-${axis}`,
				label,
				lock: false,
				value: (ctx) => String(target(ctx)?.ws.pageSetup?.[field] ?? 0),
				run: (ctx, arg) => {
					const t = target(ctx);
					if (!t || typeof arg !== 'string' || !/^\d+$/.test(arg)) return;
					const count = Number(arg);
					if (count > 32767) return;
					const setup = { fitToWidth: 0, fitToHeight: 0, ...t.ws.pageSetup, [field]: count };
					const fit = setup.fitToWidth || setup.fitToHeight;
					t.session.setPageSetup(t.sheet, {
						fitToWidth: fit ? setup.fitToWidth : undefined,
						fitToHeight: fit ? setup.fitToHeight : undefined,
					});
				},
			}),
		),
		editing({
			id: 'page.scale',
			label: 'Scale %',
			lock: false,
			value: (ctx) => String(target(ctx)?.ws.pageSetup?.scale ?? 100),
			run: (ctx, arg) => {
				const t = target(ctx);
				if (!t || typeof arg !== 'string' || !/^\d+\s*%?$/.test(arg.trim())) return;
				const scale = Number.parseInt(arg, 10);
				if (scale < 10 || scale > 400) return;
				t.session.setPageSetup(t.sheet, { scale, fitToWidth: undefined, fitToHeight: undefined });
			},
		}),
		editing({
			id: 'page.print-titles',
			label: 'Print Titles',
			icon: icon('printTitles'),
			lock: false,
			run: (ctx) => void ctx.dialogs.open('page-setup', { tab: 'sheet' }),
		}),
	];
}
