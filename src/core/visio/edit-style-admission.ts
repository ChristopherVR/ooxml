import { attribute, children } from './sheet';
import { fail } from './package-common';
import { executableCellFormula } from './cell-formula';
import {
	formattingRow,
	delegatedFormattingCell,
	formattingRows,
	mergeFormattingRows,
	type FormattingRowSources,
} from './edit-formatting-rows';
export { formattingRow } from './edit-formatting-rows';
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
function namedCell(sheet: Element, name: string): Element | undefined {
	const [section, index, cell] = name.split('.');
	if (!cell) return formattingCell(sheet, name);
	const row = formattingRow(sheet, section!, index);
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
function formattingStyles(
	shape: Element,
	document: Element,
	category: FormattingCategory,
	gate: boolean,
): Element[] {
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
	const result: Element[] = [];
	while (id !== undefined) {
		if (seen.has(id) || seen.size >= 64)
			fail('EDIT_PROTECTED_CELL', 'Formatting style ancestry is cyclic or too deep.');
		seen.add(id);
		const style = styles.get(id);
		if (!style) fail('EDIT_PROTECTED_CELL', 'Formatting style ancestry cannot be resolved.');
		// A disabled ancestor contributes no cells; enabled child overrides survive.
		// Protection still uses all categories, matching geometry admission.
		if (gate && !categoryEnabled(style, category)) break;
		result.push(style);
		const parent = attribute(style, category);
		id = parent !== id ? parent : undefined;
	}
	return result;
}

export function effectiveFormattingRows(
	shape: Element,
	document: Element,
	section: string,
): FormattingRowSources {
	let result: FormattingRowSources = new Map();
	for (const style of formattingStyles(shape, document, 'TextStyle', true).reverse())
		result = mergeFormattingRows(result, style, section, uniqueFormattingCells);
	return mergeFormattingRows(result, shape, section, uniqueFormattingCells);
}
export interface FormattingRowContext {
	source: FormattingRowSources;
	local: Map<string, Element>;
}
export function formattingRowContext(
	shape: Element,
	document: Element,
	section: string,
): FormattingRowContext {
	return {
		source: effectiveFormattingRows(shape, document, section),
		local: formattingRows(shape, section),
	};
}

export function effectiveShapeCell(
	shape: Element,
	document: Element,
	name: string,
	category: FormattingCategory,
	rows?: FormattingRowContext,
): Element | undefined {
	const [section, index, cellName] = name.split('.');
	const localRow = rows?.local.get(index!);
	const local = rows
		? localRow
			? formattingCell(localRow, cellName!)
			: undefined
		: namedCell(shape, name);
	const delegatedLocal = local && attribute(local, 'F') === 'Inh';
	if (
		delegatedLocal &&
		!(category === 'TextStyle' && cellName && ['Character', 'Paragraph'].includes(section!))
	)
		fail('EDIT_PROTECTED_CELL', 'Inherited formatting cells cannot be overwritten.');
	if (local && !cellName) return local;
	if (cellName) {
		const sources = rows?.source ?? effectiveFormattingRows(shape, document, section!);
		const cells = sources.get(index!) ?? sources.get('0');
		if (
			cells &&
			!cells.has(cellName) &&
			[...cells.keys()].some((key) => key.toLowerCase() === cellName.toLowerCase())
		)
			fail('EDIT_AMBIGUOUS_CELL', 'Noncanonical formatting cell names cannot be overridden.');
		const effective = delegatedFormattingCell(cells?.get(cellName));
		if (delegatedLocal) {
			if (!effective)
				fail(
					'EDIT_PROTECTED_CELL',
					'Inherited formatting cannot be overwritten without a proven ancestor.',
				);
			// Cacheless Inh cells contribute no parser value, but their retained unit would
			// become active after writing a local value. Prove that unit before retaining it.
			if (
				!local.hasAttribute('V') &&
				local.hasAttribute('U') &&
				visioFormulaCachedValue('0', attribute(local, 'U')).unit !==
					visioFormulaCachedValue('0', attribute(effective, 'U')).unit
			)
				fail('EDIT_PROTECTED_CELL', 'Inherited formatting has incompatible local units.');
		}
		return effective;
	}
	for (const style of formattingStyles(
		shape,
		document,
		category,
		!/^(Lock|LayerMember$|DisplayLevel$)/.test(name),
	)) {
		const cell = namedCell(style, name);
		if (cell && attribute(cell, 'F') !== 'Inh') return cell;
		if (
			cell &&
			(!attribute(style, category) || attribute(style, category) === attribute(style, 'ID'))
		)
			fail('EDIT_PROTECTED_CELL', 'Delegated formatting has no explicit ancestor.');
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

/** Native ordinary shapes explicitly cache an empty LayerMember string. */
export function assertUnlayeredShape(shape: Element, document: Element): void {
	for (const category of ['LineStyle', 'FillStyle', 'TextStyle'] as const) {
		const cell = effectiveShapeCell(shape, document, 'LayerMember', category);
		if (!cell) continue;
		if (
			attribute(cell, 'V') !== '' ||
			cell.hasAttribute('E') ||
			executableCellFormula(attribute(cell, 'F'))
		)
			fail(
				'UNSUPPORTED_FORMAT_EDIT',
				'Layered or unresolved layer membership is not yet supported.',
			);
	}
}

export { assertEditableFormattingCell } from './edit-formatting-cell-admission';
