import type { CellView, EdgeView } from 'ooxml-core/xlsx';
export const escapeHtml = (text: string) =>
	text.replace(/[&<>"']/g, (ch) => `&#${ch.charCodeAt(0)};`);

const edge = (view: EdgeView | undefined) =>
	view ? `${view.style === 'double' ? 3 : view.widthPx}px ${view.style} ${view.color}` : 'none';

/** Inline CSS for one cell view. */
export function cellCss(view: CellView): string {
	const { font } = view;
	const parts = [
		`font-family:"${font.family.replace(/["\\]/g, '')}",Calibri,sans-serif`,
		`font-size:${font.sizePx}px`,
		`color:${font.color}`,
		`text-align:${view.hAlign === 'centerContinuous' ? 'center' : view.hAlign === 'distributed' || view.hAlign === 'fill' ? 'justify' : view.hAlign}`,
		`vertical-align:${view.vAlign === 'center' ? 'middle' : view.vAlign === 'top' ? 'top' : 'bottom'}`,
		`white-space:${view.wrap ? 'pre-wrap' : 'pre'}`,
		`border-top:${edge(view.borders.top)}`,
		`border-right:${edge(view.borders.right)}`,
		`border-bottom:${edge(view.borders.bottom)}`,
		`border-left:${edge(view.borders.left)}`,
	];
	if (font.bold) parts.push('font-weight:700');
	if (font.italic) parts.push('font-style:italic');
	const lines = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(
		Boolean,
	);
	if (lines.length) parts.push(`text-decoration:${lines.join(' ')}`);
	if (view.indentPx) parts.push(`padding-left:${view.indentPx + 2}px`);
	const fill = view.fill;
	if (fill && 'background' in fill) parts.push(`background:${fill.background}`);
	else if (fill && 'gradient' in fill) parts.push(`background:${fill.gradient}`);
	else if (fill && 'pattern' in fill) parts.push(`background:${fill.fg}`);
	return parts.join(';');
}
