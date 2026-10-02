import { elements } from '../xml/index.js';
import type { VisioDiagnostic, VisioImage } from './model.js';
import { metadata, metadataAttributes } from './metadata.js';

export const VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main';
export const VISIO_LEGACY_NS = 'http://schemas.microsoft.com/office/visio/2011/1/core';
export const children = (node: Element | undefined, name: string): Element[] =>
	node
		? elements(node).filter(
				(child) =>
					child.localName === name &&
					(child.namespaceURI === VISIO_NS || child.namespaceURI === VISIO_LEGACY_NS),
			)
		: [];
export const child = (node: Element | undefined, name: string) => children(node, name)[0];
export const attribute = (node: Element | undefined, name: string): string | undefined =>
	node?.getAttribute(name) ?? undefined;
export const yes = (value: string | undefined) => value === '1' || value === 'true';
export interface Cell {
	value?: string;
	formula?: string;
	unit?: string;
	error?: string;
}
export type Cells = Map<string, Cell>;
export interface Row {
	index: string;
	name?: string;
	type: string;
	cells: Cells;
	deleted: boolean;
}
export interface Section {
	name: string;
	index: string;
	cells: Cells;
	rows: Map<string, Row>;
	deleted: boolean;
}
export interface Sheet {
	cells: Cells;
	sections: Map<string, Section>;
}
export interface RawShape extends Sheet {
	id: string;
	attributes: Map<string, string>;
	text?: Element;
	children: RawShape[];
	foreign: boolean;
	foreignData?: Element;
	image?: VisioImage;
	deleted: boolean;
}
export type Report = (code: string, message: string, context?: Partial<VisioDiagnostic>) => void;
export function readCells(node: Element): Cells {
	return new Map(
		children(node, 'Cell').map((cell) => {
			const value = attribute(cell, 'V');
			const formula = attribute(cell, 'F');
			const unit = attribute(cell, 'U');
			const error = attribute(cell, 'E');
			if (unit !== undefined) metadata(unit, 128, 'Cell unit');
			if (error !== undefined) metadata(error, 1024, 'Cell error');
			const name = metadata(attribute(cell, 'N') ?? '', 256, 'Cell name');
			if (value !== undefined && ['Font', 'BulletFont'].includes(name))
				metadata(value, 1024, 'Font name');
			if (value !== undefined && ['Name', 'NameUniv', 'BulletStr'].includes(name))
				metadata(value, 4096, 'Display label');
			if (value !== undefined && name === 'LayerMember')
				metadata(value, 8192, 'Cached layer membership');
			return [
				name,
				{
					...(value === undefined ? {} : { value }),
					...(formula === undefined ? {} : { formula }),
					...(unit === undefined ? {} : { unit }),
					...(error === undefined ? {} : { error }),
				},
			];
		}),
	);
}
export function readSheet(node: Element | undefined): Sheet {
	if (!node) return { cells: new Map(), sections: new Map() };
	const sections = new Map<string, Section>();
	for (const section of children(node, 'Section')) {
		const name = metadata(attribute(section, 'N') ?? '', 256, 'Section name');
		const index = metadata(attribute(section, 'IX') ?? '0', 256, 'Section index');
		const rows = new Map<string, Row>();
		for (const row of children(section, 'Row')) {
			const name = attribute(row, 'N');
			if (name !== undefined) metadata(name, 256, 'Row name');
			const index = metadata(attribute(row, 'IX') ?? attribute(row, 'N') ?? '0', 256, 'Row index');
			rows.set(index, {
				index,
				...(name === undefined ? {} : { name }),
				type: metadata(attribute(row, 'T') ?? '', 128, 'Geometry row type'),
				cells: readCells(row),
				deleted: yes(attribute(row, 'Del')),
			});
		}
		sections.set(`${name}:${index}`, {
			name,
			index,
			cells: readCells(section),
			rows,
			deleted: yes(attribute(section, 'Del')),
		});
	}
	return { cells: readCells(node), sections };
}
export function readShapes(node: Element | undefined): RawShape[] {
	return children(child(node, 'Shapes'), 'Shape').map((shape) => {
		const text = child(shape, 'Text');
		const foreignData = child(shape, 'ForeignData');
		return {
			...readSheet(shape),
			id: metadata(attribute(shape, 'ID') ?? '', 256, 'Shape ID'),
			attributes: metadataAttributes(shape),
			...(text ? { text } : {}),
			children: readShapes(shape),
			foreign: !!foreignData,
			...(foreignData ? { foreignData } : {}),
			deleted: yes(attribute(shape, 'Del')),
		};
	});
}
export function mergeCells(base: Cells, local: Cells): Cells {
	const result = new Map(base);
	for (const [key, value] of local) {
		if (value.formula === 'Inh' && value.value === undefined && value.error === undefined) continue;
		const unit = value.unit ?? base.get(key)?.unit;
		result.set(key, { ...value, ...(unit === undefined ? {} : { unit }) });
	}
	return result;
}
export function mergeSheets(base: Sheet, local: Sheet): Sheet {
	const sections = new Map(base.sections);
	for (const [key, section] of local.sections) {
		if (section.deleted) {
			sections.delete(key);
			continue;
		}
		const previous = sections.get(key);
		const rows = new Map(previous?.rows);
		for (const [index, row] of section.rows) {
			if (row.deleted) {
				rows.delete(index);
				continue;
			}
			const old =
				previous?.rows.get(index) ??
				(['Character', 'Paragraph', 'Tabs'].includes(section.name)
					? previous?.rows.get('0')
					: undefined);
			rows.set(index, {
				...row,
				...(row.name === undefined && old?.name !== undefined ? { name: old.name } : {}),
				type: row.type || old?.type || '',
				cells: mergeCells(old?.cells ?? new Map(), row.cells),
			});
		}
		sections.set(key, {
			...section,
			cells: mergeCells(previous?.cells ?? new Map(), section.cells),
			rows,
		});
	}
	return { cells: mergeCells(base.cells, local.cells), sections };
}
/** Only numeric cached results are used. Formula text is never evaluated. */
export function number(cells: Cells, name: string, fallback: number, report?: Report): number {
	const cell = cells.get(name);
	if (!cell) return fallback;
	const raw = cell.value?.trim();
	if (raw && raw.length <= 128 && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(raw)) {
		const value = Number(raw);
		if (Number.isFinite(value) && Math.abs(value) <= 1e9) return value;
	}
	report?.(
		'missing-cached-value',
		`Cell ${name} has no usable cached numeric value; a default was used.`,
	);
	return fallback;
}
export const sectionRows = (sheet: Sheet, name: string) =>
	[...sheet.sections.values()]
		.filter((s) => s.name === name && !s.deleted)
		.flatMap((s) => [...s.rows.values()].filter((r) => !r.deleted));
export const emptySheet = (): Sheet => ({ cells: new Map(), sections: new Map() });
/** E is the current formula error; V remains the last valid saved value. */
export function reportCachedErrors(sheet: Sheet, report: Report): void {
	const visit = (cells: Cells) => {
		for (const [name, cell] of cells) {
			if (cell.error !== undefined)
				report(
					'cached-cell-error',
					`Cell ${name} records a formula error; only its last valid cached value is available.`,
				);
		}
	};
	visit(sheet.cells);
	for (const section of sheet.sections.values()) {
		visit(section.cells);
		for (const row of section.rows.values()) visit(row.cells);
	}
}
