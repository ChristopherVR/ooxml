/** A section's `w:sectPr`: page size, margins, the header and footer references and the columns. */
import { el } from '../load/legacy-doc/xml';
import type { ColumnLayout } from './flow';
import type { SectionLayout } from './package';

const twips = (points: number) => Math.round(points * 20);

export function sectionXml(section: SectionLayout, references: string): string {
	const landscape = section.width > section.height;
	return (
		`<w:sectPr>${references}` +
		(section.continuous ? el('w:type', { 'w:val': 'continuous' }) : '') +
		el('w:pgSz', {
			'w:w': twips(section.width),
			'w:h': twips(section.height),
			'w:orient': landscape ? 'landscape' : undefined,
		}) +
		el('w:pgMar', {
			'w:top': twips(section.margins.top),
			'w:right': twips(section.margins.right),
			'w:bottom': twips(section.margins.bottom),
			'w:left': twips(section.margins.left),
			'w:header': twips(section.header),
			'w:footer': twips(section.footer),
			'w:gutter': 0,
		}) +
		columnsXml(section.columns) +
		'</w:sectPr>'
	);
}

function columnsXml(columns: ColumnLayout | undefined): string {
	if (!columns) return '';
	const space = twips(columns.space);
	if (!columns.widths) return el('w:cols', { 'w:num': columns.count, 'w:space': space });
	return el(
		'w:cols',
		{ 'w:num': columns.count, 'w:space': space, 'w:equalWidth': 0 },
		columns.widths
			.map((width, index) =>
				el('w:col', {
					'w:w': twips(width),
					'w:space': index < columns.count - 1 ? space : undefined,
				}),
			)
			.join(''),
	);
}
