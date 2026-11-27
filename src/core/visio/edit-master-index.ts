import { attribute, children } from './sheet';
import { fail } from './package-common';
import { visioFormulaCachedValue } from './formula';
import { shapeDataFormulaSection } from './cell-formula';

function inheritedProtection(node: Element, ancestor: Element | undefined): void {
	try {
		const boolean = (cell: Element | undefined) => {
			if (!cell || cell.hasAttribute('E'))
				fail('EDIT_PROTECTED_CELL', 'Inherited protection has an unresolved or error cache.');
			const value = visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'));
			if (value.unit !== 'scalar' || ![0, 1].includes(value.value))
				fail('EDIT_PROTECTED_CELL', 'Inherited protection requires scalar boolean caches.');
			return value.value;
		};
		if (boolean(node) !== boolean(ancestor))
			fail(
				'EDIT_PROTECTED_CELL',
				'Delegated master style protection cache disagrees with its ancestor.',
			);
	} catch (error) {
		fail(
			'EDIT_PROTECTED_CELL',
			`Inherited master style protection cannot be proven: ${error instanceof Error ? error.message : 'invalid cache'}`,
		);
	}
}

export function masterCells(
	sheet: Element | undefined,
	charge: () => void = () => {},
): Map<string, Element> {
	const result = new Map<string, Element>();
	if (!sheet) return result;
	const add = (name: string, node: Element) => {
		charge();
		const id = name.toLowerCase();
		if (result.has(id) && result.get(id) !== node)
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Ambiguous effective master cells.');
		result.set(id, node);
	};
	for (const node of children(sheet, 'Cell')) add(attribute(node, 'N') ?? '', node);
	for (const section of children(sheet, 'Section')) {
		const name = shapeDataFormulaSection(attribute(section, 'N') ?? '');
		const prefix =
			name === 'Geometry'
				? `Geometry${Number(attribute(section, 'IX') ?? '0') + 1}`
				: name === 'Control'
					? 'Controls'
					: name;
		for (const node of children(section, 'Cell'))
			add(`${prefix}.${attribute(node, 'N') ?? ''}`, node);
		for (const row of children(section, 'Row'))
			for (const node of children(row, 'Cell')) {
				const n = attribute(node, 'N') ?? '';
				const rowName = attribute(row, 'N');
				const index =
					attribute(row, 'IX') ??
					(prefix === 'Controls' && /^Row_\d+$/.test(rowName ?? '')
						? String(Number(rowName!.slice(4)) - 1)
						: undefined);
				const ix = index ?? '0',
					r = rowName ?? (prefix === 'Controls' ? `Row_${Number(ix) + 1}` : ix);
				// Formulas count Scratch, Controls and Connections rows from one (Scratch.Y1 is the
				// row with IX 0) and say Connections for the Connection section; geometry rows
				// already carry their one-based index.
				const numbered = ['Controls', 'Scratch', 'Connection'].includes(prefix);
				add(
					name === 'Geometry'
						? `${prefix}.${n}${ix}`
						: numbered
							? `${prefix === 'Connection' ? 'Connections' : prefix}.${n}${Number(ix) + 1}`
							: n === 'Value' && ['User', 'Prop'].includes(name)
								? `${prefix}.${r}`
								: `${prefix}.${r}.${n}`,
					node,
				);
				if (name !== 'Geometry') add(`${prefix}.${r}.${n}`, node);
				// Formulas (Visio's point glue among them) name connection points from one.
				if (prefix === 'Connection' && index !== undefined)
					add(`Connections.${n}${Number(ix) + 1}`, node);
				if (prefix === 'Controls' && n === 'X') add(`${prefix}.${r}`, node);
				if (name !== 'Geometry' && index !== undefined) add(`${prefix}.$row${ix}.${n}`, node);
			}
	}
	return result;
}
export function masterShapes(root: Element): Element[] {
	const result: Element[] = [],
		pending = [root];
	while (pending.length) {
		const parent = pending.pop()!;
		for (const container of children(parent, 'Shapes'))
			for (const shape of children(container, 'Shape')) {
				result.push(shape);
				pending.push(shape);
			}
	}
	return result;
}
export function assertMasterMarkup(
	root: Element,
	known: ReadonlySet<Element>,
	check: () => void,
): void {
	for (const node of [root, ...Array.from(root.getElementsByTagName('*'))]) {
		check();
		if (node.namespaceURI === root.namespaceURI && node.hasAttribute('F') && !known.has(node))
			fail(
				'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
				'Master formula outside admitted sheet cells has unknown scope.',
			);
	}
}
const STYLE_CATEGORIES = ['LineStyle', 'FillStyle', 'TextStyle'] as const;
interface StyleChain {
	cells: ReadonlyMap<string, Element>;
	depth: number;
}
/**
 * Resolves the effective style cells of master sheets. The style index, each style's ancestry and
 * each line, fill and text style combination are built and charged once per proof, so an instance
 * pays only for what it adds: many instances on one style cost the style work of one.
 */
export function masterStyleResolver(
	document: Element,
	check: () => void,
	charge: () => void = () => {},
): (
	template: Element,
	instance: Element | undefined,
	verifyMovementProtection?: boolean,
) => ReadonlyMap<string, Element> {
	const styles = new Map<string, Element>();
	for (const container of children(document, 'StyleSheets'))
		for (const style of children(container, 'StyleSheet')) {
			charge();
			const id = attribute(style, 'ID');
			if (!id || styles.has(id))
				fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Ambiguous master style IDs.');
			styles.set(id, style);
		}
	const defaults = children(document, 'DocumentSheet')[0];
	const chains = new Map<string, StyleChain>();
	const proven = new Set<Element>();
	const collect = (
		id: string,
		category: string,
		verify: boolean,
		active = new Set<string>(),
	): StyleChain => {
		check();
		const key = JSON.stringify([id, category, verify]);
		const known = chains.get(key);
		if (known) return known;
		if (active.has(id))
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Invalid inherited master style ancestry.');
		active.add(id);
		const style = styles.get(id);
		if (!style) fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Missing inherited master style.');
		const own = masterCells(style, charge);
		if (!proven.has(style)) {
			assertMasterMarkup(style, new Set(own.values()), check);
			proven.add(style);
		}
		const parent = attribute(style, category);
		const inherited =
			parent !== undefined && parent !== id ? collect(parent, category, verify, active) : undefined;
		const depth = (inherited?.depth ?? 0) + 1;
		if (depth > 64)
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Invalid inherited master style ancestry.');
		const cells = new Map(inherited?.cells);
		for (const [name, node] of own) {
			charge();
			if (attribute(node, 'F') !== 'Inh') cells.set(name, node);
			else if (verify && /^lock(move[xy]|width|height|aspect|delete)$/.test(name)) {
				charge();
				inheritedProtection(node, cells.get(name));
			}
		}
		const chain = { cells, depth };
		chains.set(key, chain);
		return chain;
	};
	const combined = new Map<string, ReadonlyMap<string, Element>>();
	return (template, instance, verifyMovementProtection = false) => {
		const ids = STYLE_CATEGORIES.map(
			(category) =>
				attribute(instance, category) ??
				attribute(template, category) ??
				attribute(defaults, category) ??
				(styles.has('0') ? '0' : undefined),
		);
		const key = JSON.stringify([ids, verifyMovementProtection]);
		const known = combined.get(key);
		if (known) return known;
		const result = new Map<string, Element>();
		STYLE_CATEGORIES.forEach((category, index) => {
			const id = ids[index];
			if (id === undefined) return;
			for (const [name, node] of collect(id, category, verifyMovementProtection).cells) {
				charge();
				if (result.has(name) && result.get(name) !== node)
					fail(
						'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
						'Master style categories have ambiguous overlapping inputs.',
					);
				result.set(name, node);
			}
		});
		combined.set(key, result);
		return result;
	};
}

export interface MasterCellSource {
	node: Element;
	kind: 'template' | 'instance' | 'style';
}
/** The names each source cell is known by in `cells`, for `overlay`. */
export function masterAliases(
	cells: ReadonlyMap<string, MasterCellSource>,
	charge: () => void,
): Map<Element, ReadonlySet<string>> {
	const aliases = new Map<Element, Set<string>>();
	for (const [name, source] of cells) {
		charge();
		const names = aliases.get(source.node) ?? new Set<string>();
		names.add(name);
		aliases.set(source.node, names);
	}
	return aliases;
}
/**
 * Lays `additions` over `cells`. `aliases` must describe `cells` (`masterAliases`); it is updated
 * in place, and its name sets are replaced rather than changed, so a copy of a map shared between
 * instances of one master can be passed.
 */
export function overlay(
	cells: Map<string, MasterCellSource>,
	additions: ReadonlyMap<string, Element>,
	kind: MasterCellSource['kind'],
	charge: () => void,
	aliases: Map<Element, ReadonlySet<string>> = masterAliases(cells, charge),
): void {
	const inherited = new Map<Element, MasterCellSource>();
	const alias = (node: Element, names: readonly string[]) =>
		aliases.set(node, new Set([...(aliases.get(node) ?? []), ...names]));
	for (const [name, node] of additions) {
		charge();
		const source = cells.get(name);
		if (source) inherited.set(node, source);
	}
	for (const [name, node] of additions) {
		charge();
		if (attribute(node, 'F') === 'Inh') {
			const source = cells.get(name) ?? inherited.get(node);
			if (!source)
				fail(
					'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
					'Inherited master cell has no effective source.',
				);
			cells.set(name, source);
			alias(source.node, [name]);
			continue;
		}
		const previous = cells.get(name);
		const names = [name];
		if (previous && previous.node !== node) {
			for (const other of aliases.get(previous.node) ?? []) {
				charge();
				cells.set(other, { node, kind });
				names.push(other);
			}
			aliases.delete(previous.node);
		}
		cells.set(name, { node, kind });
		alias(node, names);
	}
}
