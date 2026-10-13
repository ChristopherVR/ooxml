import {
	pageHeaderFooter,
	renderHeaderFooterRuns,
	type HeaderFooterContext,
	type PrintLayout,
	type Workbook,
} from 'ooxml-core/xlsx';
import { escapeHtml } from './print-cell';

export interface PrintHeaderOptions {
	date?: Date;
	locale?: string;
	fileName?: string;
}

export function printPageHeaders(
	workbook: Workbook,
	sheetIndex: number,
	layout: PrintLayout,
	page: number,
	title: string,
	options: PrintHeaderOptions,
): string {
	const sheet = workbook.sheets[sheetIndex]!;
	const setup = sheet.pageSetup ?? {};
	const number = (setup.firstPageNumber ?? 1) + page - 1;
	const text = pageHeaderFooter(setup, page, number);
	const context: HeaderFooterContext = {
		page: number,
		pages: layout.pages.length,
		sheet: sheet.name,
		file: (options.fileName ?? title).split(/[\\/]/).at(-1)!,
		...(options.date ? { date: options.date } : {}),
		...(options.locale ? { locale: options.locale } : {}),
	};
	const left = (setup.alignWithMargins === false ? 0.7 : layout.margins.left) * 96;
	const right = (setup.alignWithMargins === false ? 0.7 : layout.margins.right) * 96;
	const size = ((11 * 96) / 72) * (setup.scaleWithDoc === false ? 1 : layout.scale);
	const band = (kind: 'header' | 'footer', value: string) => {
		if (!value) return '';
		const sections = renderHeaderFooterRuns(value, context).sections;
		const runs = (section: 'left' | 'center' | 'right') =>
			sections[section]
				.map(({ text, style }) => {
					const css = [
						style.font ? `font-family:${style.font.replace(/["\\;<>]/g, '')}` : '',
						style.size ? `font-size:${(style.size / 11) * size}px` : '',
						style.color ? `color:${style.color}` : '',
						`font-weight:${style.bold ? 700 : 400}`,
						`font-style:${style.italic ? 'italic' : 'normal'}`,
						`text-decoration:${[style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none'}`,
					]
						.filter(Boolean)
						.join(';');
					return `<span style="${escapeHtml(css)}">${escapeHtml(text)}</span>`;
				})
				.join('');
		return `<div class="xve-print-${kind}" style="position:absolute;${kind === 'header' ? 'top' : 'bottom'}:${layout.margins[kind] * 96}px;left:${left}px;right:${right}px;font-size:${size}px;display:flex;line-height:1.2">${(['left', 'center', 'right'] as const).map((section) => `<span style="width:33.333%;white-space:pre-wrap;text-align:${section};overflow:hidden">${runs(section)}</span>`).join('')}</div>`;
	};
	return band('header', text.header) + band('footer', text.footer);
}
