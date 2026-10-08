import { attribute, children } from './sheet';
import { fail } from './package-common';
import { editableCell } from './edit-geometry-admission';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula';

export type FormattingCategory = 'LineStyle' | 'FillStyle' | 'TextStyle';
export function uniqueFormattingCells(sheet: Element): Map<string, Element> {
	const result = new Map<string, Element>();
	const seen = new Set<string>();
	for (const cell of children(sheet, 'Cell')) {
		const name = attribute(cell, 'N');
		if (!name || seen.has(name.toLowerCase()))
			fail('EDIT_AMBIGUOUS_CELL', 'Formatting requires unique named cells.');
		result.set(name, cell);
		seen.add(name.toLowerCase());
	}
	return result;
}
export function formattingCell(sheet: Element, name: string): Element | undefined {
	const cells = uniqueFormattingCells(sheet);
	if (!cells.has(name) && [...cells.keys()].some((key) => key.toLowerCase() === name.toLowerCase()))
		fail('EDIT_AMBIGUOUS_CELL', 'Noncanonical formatting cell names cannot be overridden.');
	return cells.get(name);
}
export function formattingRow(sheet: Element, sectionName: string): Element | undefined {
	const sections = children(sheet, 'Section').filter(
		(section) => attribute(section, 'N') === sectionName,
	);
	if (sections.length > 1) fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting section is ambiguous.');
	const section = sections[0];
	if (!section) return undefined;
	const rows = children(section, 'Row');
	if (
		section.hasAttribute('Del') ||
		rows.length > 1 ||
		rows.some(
			(row) =>
				row.hasAttribute('Del') || (attribute(row, 'IX') ?? '0') !== '0' || row.hasAttribute('N'),
		)
	)
		fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting requires one live row with index zero.');
	return rows[0];
}

function namedCell(sheet: Element, name: string): Element | undefined {
	const [section, , cell] = name.split('.');
	if (!cell) return formattingCell(sheet, name);
	const row = formattingRow(sheet, section!);
	return row ? formattingCell(row, cell) : undefined;
}
function categoryEnabled(style: Element, category: FormattingCategory): boolean {
	const name =
		category === 'TextStyle'
			? 'EnableTextProps'
			: category === 'FillStyle'
				? 'EnableFillProps'
				: 'EnableLineProps';
	const cell = formattingCell(style, name);
	if (!cell) return true;
	if (cell.hasAttribute('E') || attribute(cell, 'F') === 'Inh')
		fail('EDIT_PROTECTED_CELL', 'Style category enable state cannot be proven.');
	const cache = visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'));
	if (cache.unit !== 'scalar' || ![0, 1].includes(cache.value))
		fail('EDIT_PROTECTED_CELL', 'Style category enable state must be a scalar boolean.');
	const source = executableCellFormula(attribute(cell, 'F'));
	if (source) {
		const analysis = analyzeVisioFormula(source);
		if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
			fail('EDIT_PROTECTED_CELL', 'Style category enable formula cannot be resolved safely.');
		const result = evaluateVisioFormula(source, () =>
			fail('EDIT_PROTECTED_CELL', 'Style category enable dependencies are unsupported.'),
		);
		if (result.unit !== 'scalar' || result.value !== cache.value)
			fail('EDIT_PROTECTED_CELL', 'Style category enable formula cache is stale.');
	}
	return cache.value !== 0;
}
/** Resolve source cells without trusting delegated caches or guessing style ancestry. */
export function effectiveShapeCell(
	shape: Element,
	document: Element,
	name: string,
	category: FormattingCategory,
): Element | undefined {
	const local = namedCell(shape, name);
	if (local && attribute(local, 'F') !== 'Inh') return local;
	if (local) fail('EDIT_PROTECTED_CELL', 'Inherited formatting cells cannot be overwritten.');
	const styles = new Map<string, Element>();
	for (const container of children(document, 'StyleSheets'))
		for (const style of children(container, 'StyleSheet')) {
			const id = attribute(style, 'ID');
			if (!id || styles.has(id)) fail('EDIT_PROTECTED_CELL', 'Ambiguous formatting style IDs.');
			styles.set(id, style);
		}
	const sheets = children(document, 'DocumentSheet');
	if (sheets.length > 1) fail('EDIT_PROTECTED_CELL', 'Ambiguous document formatting defaults.');
	let id =
		attribute(shape, category) ??
		attribute(sheets[0], category) ??
		(styles.has('0') ? '0' : undefined);
	const seen = new Set<string>();
	while (id !== undefined) {
		if (seen.has(id) || seen.size >= 64)
			fail('EDIT_PROTECTED_CELL', 'Formatting style ancestry is cyclic or too deep.');
		seen.add(id);
		const style = styles.get(id);
		if (!style) fail('EDIT_PROTECTED_CELL', 'Formatting style ancestry cannot be resolved.');
		// Parser category gates suppress the entire selected ancestry. Protection still
		// uses all categories, matching the conservative geometry admission contract.
		if (!/^(Lock|LayerMember$|DisplayLevel$)/.test(name) && !categoryEnabled(style, category))
			return undefined;
		const cell = namedCell(style, name),
			parent = attribute(style, category);
		if (cell && attribute(cell, 'F') !== 'Inh') return cell;
		if (cell && (parent === undefined || parent === id))
			fail('EDIT_PROTECTED_CELL', 'Delegated formatting has no explicit ancestor.');
		id = parent !== id ? parent : undefined;
	}
	return undefined;
}

export function assertShapeLocks(
	shape: Element,
	document: Element,
	locks: readonly string[],
): void {
	const inheritedShape = shape.cloneNode(true) as Element;
	for (const cell of children(inheritedShape, 'Cell'))
		if (locks.includes(attribute(cell, 'N') ?? '')) inheritedShape.removeChild(cell);
	for (const lock of locks)
		for (const [target, category] of [
			[shape, 'LineStyle'],
			[inheritedShape, 'LineStyle'],
			[inheritedShape, 'FillStyle'],
			[inheritedShape, 'TextStyle'],
		] as const) {
			const cell = effectiveShapeCell(target, document, lock, category);
			if (!cell) continue;
			if (cell.hasAttribute('E'))
				fail('EDIT_PROTECTED_CELL', 'Formatting protection has an error.');
			const cached = visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'));
			if (cached.unit !== 'scalar' || cached.value !== 0)
				fail('EDIT_PROTECTED_CELL', 'Formatting protection is active or invalid.');
			const formula = executableCellFormula(attribute(cell, 'F'));
			if (!formula) continue;
			const analysis = analyzeVisioFormula(formula);
			if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
				fail('EDIT_PROTECTED_CELL', 'Formatting protection formula cannot be proven safe.');
			const value = evaluateVisioFormula(formula, () =>
				fail('EDIT_PROTECTED_CELL', 'Formatting lock dependency is unsupported.'),
			);
			if (value.unit !== 'scalar' || value.value !== cached.value)
				fail('EDIT_PROTECTED_CELL', 'Formatting protection cache is stale.');
		}
}

/** Literal native font/color lookups are safe to replace after GUARD and reference analysis. */
export function assertEditableFormattingCell(cell: Element | undefined): void {
	const source = executableCellFormula(attribute(cell, 'F'));
	if (
		source &&
		/^\s*=?\s*(?:FONT\(\s*"[^"\r\n]*"\s*\)|RGB\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\))\s*$/i.test(source)
	) {
		if (cell?.hasAttribute('E'))
			fail('EDIT_PROTECTED_CELL', 'Cannot overwrite an error formatting cell.');
		return;
	}
	editableCell(cell);
	if (source && cell) {
		const actual = evaluateVisioFormula(source, () =>
			fail('EDIT_PROTECTED_CELL', 'Formatting dependencies cannot be overwritten.'),
		);
		const cache = visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'));
		if (
			actual.value !== cache.value ||
			(actual.unit !== cache.unit &&
				!(cache.unit === 'scalar' && actual.unit === 'length' && !cell.hasAttribute('U')))
		)
			fail('EDIT_PROTECTED_CELL', 'Formatting formula cache is stale or has incompatible units.');
	}
}
