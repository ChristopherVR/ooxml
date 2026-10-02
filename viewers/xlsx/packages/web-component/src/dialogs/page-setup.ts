// Page Setup: Page, Margins, Header/Footer and Sheet tabs. OK applies every tab in one undo step.
import type { PageSetup, PrintOptions } from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { tabs } from './fields.js';
import { showDialog } from './frame.js';
import { headerTab, marginsTab, pageTab, sheetTab } from './page-setup-tabs.js';

export interface PageSetupProps {
	tab?: 'page' | 'margins' | 'header' | 'sheet';
}

export function openPageSetup(
	ctx: EditorContext,
	props: PageSetupProps = {},
): Promise<Partial<PageSetup> | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const setup = t.ws.pageSetup ?? {};
	const page = pageTab(ctx, setup);
	const margins = marginsTab(ctx, setup, t.ws);
	const header = headerTab(ctx, setup);
	const sheet = sheetTab(ctx, setup, t.ws);
	const view = tabs(
		ctx,
		[
			{ id: 'page', label: 'Page', panel: page.panel },
			{ id: 'margins', label: 'Margins', panel: margins.panel },
			{ id: 'header', label: 'Header/Footer', panel: header.panel },
			{ id: 'sheet', label: 'Sheet', panel: sheet.panel },
		],
		props.tab ?? 'page',
	);
	return showDialog<Partial<PageSetup>>(ctx, {
		name: 'page-setup',
		heading: 'Page Setup',
		wide: true,
		body: view.element,
		opened: () =>
			view.element.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.focus(),
		submit: () => {
			const patch: Partial<PageSetup> = {};
			for (const [id, tab] of [
				['page', page],
				['margins', margins],
				['header', header],
				['sheet', sheet],
			] as const) {
				const result = tab.read();
				if (typeof result === 'string') {
					view.show(id);
					ctx.toast(ctx.t(result), 'warning');
					return undefined;
				}
				Object.assign(patch, result);
			}
			t.session.batch('Page Setup', () => {
				t.session.setPageSetup(t.sheet, patch);
				const named = t.workbook.definedNames.some(
					(n) => n.name === '_xlnm.Print_Area' && n.localSheet === t.sheet,
				);
				if (!patch.printArea && named) t.session.deleteDefinedName('_xlnm.Print_Area', t.sheet);
				const options: Required<PrintOptions> = {
					gridLines: sheet.gridlines(),
					headings: sheet.headings(),
					...margins.centered(),
				};
				const current = t.ws.printOptions ?? {};
				const keys = Object.keys(options) as (keyof PrintOptions)[];
				if (keys.some((key) => !!current[key] !== options[key]))
					t.session.setPrintOptions(t.sheet, options);
			});
			return patch;
		},
	});
}
