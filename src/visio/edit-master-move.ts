import { attribute, children } from './sheet.js';
import { fail } from './package-common.js';
import type { VisioGeometryEdit } from './edit-commands.js';
import type { MasterCellSource } from './edit-master-index.js';
import { cells, editableCell, numeric } from './edit-geometry-admission.js';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula.js';
import { executableCellFormula } from './cell-formula.js';
import { indexCells, key } from './edit-recalculate-index.js';

interface MoveBinding {
	instance?: Element;
	template: Element;
	cells: ReadonlyMap<string, MasterCellSource>;
	inheritedCells: ReadonlyMap<string, MasterCellSource>;
}
const protection = [
	'LockMoveX',
	'LockMoveY',
	'LockWidth',
	'LockHeight',
	'LockAspect',
	'LockDelete',
];
const dimensions = ['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY'];
export interface MasterMoveProof {
	pins: ReadonlySet<Element>;
	dimensions: ReadonlyMap<Element, { width: number; height: number }>;
}
export const emptyMasterMoveProof = (): MasterMoveProof => ({
	pins: new Set(),
	dimensions: new Map(),
});
/** Preparing pin leaves is not authorization: the entire master dependency proof must still succeed. */
export function prepareMasterMovePins(
	roots: ReadonlyMap<string, Element>,
	commands: readonly VisioGeometryEdit[],
	bindings: readonly MoveBinding[],
	check: () => void,
): MasterMoveProof {
	const result = new Set<Element>();
	const provenDimensions = new Map<Element, { width: number; height: number }>();
	let work = 100_000;
	const charge = (amount = 1) => {
		check();
		work -= amount;
		if (work < 0)
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Master move preparation budget exceeded.');
	};
	const analyze = (formula: string) => {
		charge(formula.length);
		return analyzeVisioFormula(formula);
	};
	for (const command of commands) {
		charge();
		const root = roots.get(command.pageId);
		const instance = children(children(root, 'Shapes')[0], 'Shape').find(
			(shape) => attribute(shape, 'ID') === command.shapeId,
		);
		if (!instance?.hasAttribute('Master')) continue;
		if (command.type !== 'move-shape')
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Master-linked resizing and deletion remain unsupported.');
		const binding = bindings.find((candidate) => candidate.instance === instance);
		if (!binding) fail('UNSUPPORTED_GEOMETRY_EDIT', 'A unique effective master root is required.');
		for (const node of [instance, binding.template]) {
			charge();
			if (
				attribute(node, 'Type') !== 'Shape' ||
				children(node, 'Shapes').length ||
				children(node, 'ForeignData').length ||
				children(node, 'Rel').length ||
				['1', 'true'].includes(attribute(node, 'Del') ?? '')
			)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Master moves require top-level local 2D shape roots.');
		}
		for (const name of ['BeginX', 'BeginY', 'EndX', 'EndY'])
			if (binding.cells.has(name.toLowerCase()))
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Master lines and endpoint routing remain unsupported.');
		const oneD = binding.cells.get('oned')?.node;
		charge();
		if (
			oneD &&
			(numeric(oneD) !== 0 ||
				visioFormulaCachedValue(attribute(oneD, 'V') ?? '', attribute(oneD, 'U')).unit !==
					'scalar' ||
				executableCellFormula(attribute(oneD, 'F')))
		)
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'A static effective 2D master is required.');
		const local = cells(instance);
		for (const name of dimensions) {
			const effective = binding.cells.get(name.toLowerCase())?.node;
			for (const node of new Set([effective, ...(local.has(name) ? [local.get(name)] : [])])) {
				charge();
				numeric(node);
				if (visioFormulaCachedValue('0', attribute(node, 'U') ?? 'DL').unit !== 'length')
					fail('EDIT_FORMULA_UNIT', 'Master transform caches require length units.');
			}
		}
		const dimension = (name: string) =>
			numeric(local.get(name) ?? binding.cells.get(name.toLowerCase())?.node);
		const width = dimension('Width'),
			height = dimension('Height');
		if (width <= 0 || height <= 0)
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Positive proven effective dimensions are required.');
		provenDimensions.set(instance, { width, height });
		for (const name of protection) {
			charge();
			const node = binding.cells.get(name.toLowerCase())?.node;
			const cached = node ? numeric(node) : NaN;
			const moving = name === 'LockMoveX' || name === 'LockMoveY';
			if (
				!node ||
				visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).unit !==
					'scalar' ||
				![0, 1].includes(cached) ||
				(moving && cached !== 0)
			)
				fail(
					'EDIT_PROTECTED_CELL',
					'Effective master protection must be proven boolean with inactive movement locks.',
				);
			const formula = executableCellFormula(attribute(node, 'F'));
			if (formula) {
				const analysis = analyze(formula);
				if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
					fail('EDIT_PROTECTED_CELL', 'Effective master protection dependencies are unsupported.');
				const value = evaluateVisioFormula(
					formula,
					() => fail('EDIT_PROTECTED_CELL', 'Unexpected protection dependency.'),
					{ onStep: () => charge() },
				);
				if (value.unit !== 'scalar' || value.value !== cached)
					fail('EDIT_PROTECTED_CELL', 'Effective master protection cache is stale.');
			}
		}
		for (const name of ['PinX', 'PinY']) {
			charge();
			const node = local.get(name);
			if (!node)
				fail('EDIT_PROTECTED_CELL', 'The rotation pin must be an explicit local override.');
			if (binding.cells.get(name.toLowerCase())?.node !== node)
				fail('EDIT_PROTECTED_CELL', 'The rotation pin must be an explicit local override.');
			charge(executableCellFormula(attribute(node, 'F'))?.length ?? 0);
			editableCell(node);
			const inherited = binding.inheritedCells.get(name.toLowerCase())?.node;
			numeric(inherited);
			if (visioFormulaCachedValue('0', attribute(inherited, 'U') ?? 'DL').unit !== 'length')
				fail('EDIT_FORMULA_UNIT', 'Inherited rotation pins require length units.');
			const formula = executableCellFormula(attribute(inherited, 'F'));
			if (formula) {
				const analysis = analyze(formula);
				if (analysis.guarded || analysis.dynamic || analysis.unsupportedFunctions.length)
					fail('EDIT_PROTECTED_CELL', 'Protected or redirected master pins cannot be overridden.');
			}
			result.add(node);
		}
	}
	// A separate page cache depending on an authorized pin still needs a broader proof.
	// This first slice changes only the two local pin leaves, never dependent caches.
	if (result.size) {
		const indexed = indexCells(roots, { check });
		const pins = new Set(
			[...indexed.values()].filter((cell) => cell.node && result.has(cell.node)).map(key),
		);
		for (const cell of indexed.values())
			if ((charge(), cell.dependencies.some((dependency) => pins.has(dependency))))
				fail('EDIT_UNSUPPORTED_DEPENDENCY', 'A page cache depends on a master rotation pin.');
	}
	return { pins: result, dimensions: provenDimensions };
}
