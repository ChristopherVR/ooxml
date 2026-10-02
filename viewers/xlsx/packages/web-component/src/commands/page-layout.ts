// Page Layout tab: theme colours (read-only), margins, orientation, paper size, print area, the
// Page Setup dialog and the Sheet Options (gridlines and headings, view and print).
import type { PageSetup } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { editing, target, viewing } from './util.js';

type Margins = NonNullable<PageSetup['margins']>;

/** Excel's margin presets, in inches. */
export const MARGIN_PRESETS: ReadonlyArray<readonly [id: string, label: string, margins: Margins]> =
	[
		[
			'normal',
			'Normal',
			{ top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
		],
		['wide', 'Wide', { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 }],
		[
			'narrow',
			'Narrow',
			{ top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 },
		],
	];

/** SpreadsheetML paper size codes Excel's Size menu lists first. */
export const PAPER_SIZES: ReadonlyArray<readonly [code: number, label: string]> = [
	[1, 'Letter'],
	[5, 'Legal'],
	[7, 'Executive'],
	[9, 'A4'],
	[8, 'A3'],
	[11, 'A5'],
	[13, 'B5 (JIS)'],
	[20, 'Envelope #10'],
	[27, 'Envelope DL'],
];

const page = (ctx: EditorContext): PageSetup => target(ctx)?.ws.pageSetup ?? {};

const setPage = (ctx: EditorContext, patch: Partial<PageSetup>): void => {
	const t = target(ctx);
	if (t) t.session.setPageSetup(t.sheet, patch);
};

export function pageLayoutCommands(): Command[] {
	return [
		viewing({
			id: 'page.theme-colors',
			label: 'Colors',
			icon: icon('themeColors'),
			run: (ctx) =>
				ctx.toast(
					ctx.t(
						'Theme colors are shown as read only; changing the workbook theme is not supported yet.',
					),
					'info',
				),
		}),
		editing({
			id: 'page.margins',
			label: 'Margins',
			icon: icon('margins'),
			lock: false,
			run: (ctx, arg) => {
				const preset = MARGIN_PRESETS.find(([id]) => id === arg) ?? MARGIN_PRESETS[0];
				if (preset) setPage(ctx, { margins: { ...preset[2] } });
			},
		}),
		editing({
			id: 'page.margins-custom',
			label: 'Custom Margins...',
			icon: icon('margins'),
			lock: false,
			run: (ctx) => void ctx.dialogs.open('page-setup', { tab: 'margins' }),
		}),
		...(['portrait', 'landscape'] as const).map((orientation) =>
			editing({
				id: `page.orientation-${orientation}`,
				label: orientation === 'portrait' ? 'Portrait' : 'Landscape',
				icon: icon('pageOrientation'),
				lock: false,
				checked: (ctx) => (page(ctx).orientation ?? 'portrait') === orientation,
				run: (ctx) => setPage(ctx, { orientation }),
			}),
		),
		editing({
			id: 'page.size',
			label: 'Size',
			icon: icon('pageSize'),
			lock: false,
			run: (ctx, arg) => {
				const code = Number(arg);
				if (PAPER_SIZES.some(([c]) => c === code)) setPage(ctx, { paperSize: code });
			},
		}),
		editing({
			id: 'page.size-more',
			label: 'More Paper Sizes...',
			icon: icon('pageSize'),
			lock: false,
			run: (ctx) => void ctx.dialogs.open('page-setup', { tab: 'page' }),
		}),
		editing({
			id: 'page.print-area-set',
			label: 'Set Print Area',
			icon: icon('printArea'),
			lock: false,
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setPageSetup(t.sheet, { printArea: t.range });
			},
		}),
		editing({
			id: 'page.print-area-clear',
			label: 'Clear Print Area',
			icon: icon('printArea'),
			lock: false,
			enabled: (ctx) => !!page(ctx).printArea,
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				t.session.batch('Clear Print Area', () => {
					t.session.setPageSetup(t.sheet, { printArea: undefined });
					// A loaded workbook may also carry the name; drop it so the writer does not restore it.
					if (
						t.workbook.definedNames.some(
							(n) => n.name === '_xlnm.Print_Area' && n.localSheet === t.sheet,
						)
					)
						t.session.deleteDefinedName('_xlnm.Print_Area', t.sheet);
				});
			},
		}),
		editing({
			id: 'page.setup',
			label: 'Page Setup',
			icon: icon('pageSetup'),
			lock: false,
			run: (ctx, arg) => void ctx.dialogs.open('page-setup', arg),
		}),
		viewing({
			id: 'page.gridlines-view',
			label: 'View Gridlines',
			icon: icon('gridlines'),
			checked: (ctx) => target(ctx)?.ws.view.showGridLines !== false,
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setSheetView(t.sheet, { showGridLines: !t.ws.view.showGridLines });
			},
		}),
		viewing({
			id: 'page.headings-view',
			label: 'View Headings',
			icon: icon('headings'),
			checked: (ctx) => target(ctx)?.ws.view.showHeaders !== false,
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setSheetView(t.sheet, { showHeaders: !t.ws.view.showHeaders });
			},
		}),
		...(['gridLines', 'headings'] as const).map((name) =>
			editing({
				id: name === 'gridLines' ? 'page.gridlines-print' : 'page.headings-print',
				label: name === 'gridLines' ? 'Print Gridlines' : 'Print Headings',
				icon: icon(name === 'gridLines' ? 'gridlines' : 'headings'),
				lock: false,
				checked: (ctx) => {
					const t = target(ctx);
					return !!t?.ws.printOptions?.[name];
				},
				run: (ctx) => {
					const t = target(ctx);
					if (t) t.session.setPrintOptions(t.sheet, { [name]: !t.ws.printOptions?.[name] });
				},
			}),
		),
	];
}
