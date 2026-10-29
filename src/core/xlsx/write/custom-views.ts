// Writes `customWorkbookViews` and a sheet's `customSheetViews`. The sheet-level view goes through
// the writers of the sheet itself (a scratch worksheet carries the recorded state), so panes,
// page setup, print options and the auto-filter are written exactly as the sheet would write them.
import { elements, parseXml } from '../../xml/index';
import { createWorksheet } from '../workbook';
import type { CustomSheetView, CustomWorkbookView } from '../model-custom-views';
import { printOptionsXml } from '../read/print-options';
import { outerXml } from '../read/xml-util';
import { pageXml } from './page-setup';
import { escapeAttr } from './xml-out';
import { autoFilterXml, sheetViewsXml } from './sheet-props';

const attr = (name: string, value: string | number | undefined) =>
	value === undefined ? '' : ` ${name}="${escapeAttr(String(value))}"`;

export function customWorkbookViewsXml(views: readonly CustomWorkbookView[] | undefined): string {
	if (!views?.length) return '';
	const items = views
		.map((view) => {
			// The schema requires the window size; Excel's own default fills it in when missing.
			const withWindow = { windowWidth: '28800', windowHeight: '15370', ...view.attrs };
			const rest = Object.entries(withWindow)
				.map(([name, value]) => attr(name, value))
				.join('');
			return (
				`<customWorkbookView${attr('name', view.name)}${attr('guid', view.guid)}` +
				`${view.includePrintSettings ? '' : ' includePrintSettings="0"'}` +
				`${view.includeHiddenRowCol ? '' : ' includeHiddenRowCol="0"'}${rest}/>`
			);
		})
		.join('');
	return `<customWorkbookViews>${items}</customWorkbookViews>`;
}

/** The attributes and pane/selection children the sheet writer produces for a view. */
function viewParts(view: CustomSheetView): { attributes: string; children: string } {
	const scratch = createWorksheet('view', 1);
	scratch.view = view.view;
	const root = parseXml(sheetViewsXml(scratch, false)).documentElement;
	const node = elements(root)[0];
	if (!node) return { attributes: '', children: '' };
	const get = (name: string) => node.getAttribute(name) ?? undefined;
	const attributes =
		attr('scale', get('zoomScale')) +
		attr('showGridLines', get('showGridLines')) +
		attr('showRowCol', get('showRowColHeaders')) +
		attr('zeroValues', get('showZeros')) +
		attr('topLeftCell', get('topLeftCell')) +
		attr('view', get('view')) +
		attr('showFormulas', get('showFormulas'));
	const children = elements(node)
		.map((child) => outerXml(child).replace(/ xmlns="[^"]*"/, ''))
		.join('');
	return { attributes, children };
}

export function customSheetViewsXml(views: readonly CustomSheetView[] | undefined): string {
	if (!views?.length) return '';
	const items = views
		.map((view) => {
			const scratch = createWorksheet('view', 1);
			if (view.pageSetup) scratch.pageSetup = view.pageSetup;
			if (view.filter) scratch.autoFilter = view.filter;
			const { attributes, children } = viewParts(view);
			const page = pageXml(scratch);
			const rest = Object.entries(view.attrs)
				.map(([name, value]) => attr(name, value))
				.join('');
			const fit = view.pageSetup?.fitToWidth !== undefined ? ' fitToPage="1"' : '';
			const flags =
				(view.hiddenRows ? ' hiddenRows="1"' : '') +
				(view.hiddenColumns ? ' hiddenColumns="1"' : '') +
				(view.filter ? ' filter="1"' : '') +
				(view.showAutoFilter ? ' showAutoFilter="1"' : '');
			const [rowBreaks, colBreaks, others] = [
				view.extra.filter((x) => x.startsWith('<rowBreaks')),
				view.extra.filter((x) => x.startsWith('<colBreaks')),
				view.extra.filter((x) => !x.startsWith('<rowBreaks') && !x.startsWith('<colBreaks')),
			];
			return (
				`<customSheetView${attr('guid', view.guid)}${attributes}${fit}${flags}${rest}>` +
				`${children}${rowBreaks.join('')}${colBreaks.join('')}${page.margins}` +
				`${printOptionsXml(view.printOptions)}${page.setup}${page.headerFooter}` +
				`${view.filter ? autoFilterXml(scratch) : ''}${others.join('')}</customSheetView>`
			);
		})
		.join('');
	return `<customSheetViews>${items}</customSheetViews>`;
}
