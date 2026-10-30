import { attribute, children, VISIO_NS } from './sheet';
import { evaluateVisioFormula, parseVisioFormula, visioFormulaCachedValue } from './formula';
import type { VisioFormulaReference, VisioFormulaValue } from './formula';

/**
 * Recalculation of the cached values of a master instance, for the shapes Visio's cross-functional
 * flowchart masters make (frame, lists, lanes, separators). A page stores an instance's cells as
 * `V` with `F="Inh"` where the master gives the formula, and leaves out the cells whose value is
 * the master's own. When the group's size or a list's direction changes, those caches change with
 * the master's formulas; this evaluates them (the ShapeSheet subset the editor's interpreter runs)
 * in the instance's context and writes what differs. A cell whose formula it cannot run (theme
 * lookups, `CONTAINERSHEETREF`, text measures) keeps its cache.
 */

/** A shape of an instance paired with the master shape it inherits from. */
interface Pair {
	inst: Element;
	master: Element;
	cells: Map<string, Element>;
	masterCells: Map<string, Element>;
	/** The top-level shape of this instance, whose master numbers the `Sheet.N` of its formulas. */
	group: Group;
}
interface Group {
	byMasterShape: Map<string, Pair>;
}
export interface CffMasters {
	/** Master id to the master's top-level shape. */
	get(id: string): Element | undefined;
}

/** The name a cell goes by in formulas: `Width`, `User.HeadingPos`, `Geometry1.X2`. */
function cellNames(shape: Element): Map<string, Element> {
	const names = new Map<string, Element>();
	for (const cell of children(shape, 'Cell')) names.set(attribute(cell, 'N') ?? '', cell);
	for (const section of children(shape, 'Section')) {
		const kind = attribute(section, 'N');
		if (kind !== 'User' && kind !== 'Geometry' && kind !== 'Scratch' && kind !== 'Connection')
			continue;
		const index = Number(attribute(section, 'IX') ?? 0);
		for (const cell of children(section, 'Cell'))
			if (kind === 'Geometry') names.set(`Geometry${index + 1}.${attribute(cell, 'N')}`, cell);
		for (const row of children(section, 'Row')) {
			const rowName = attribute(row, 'N');
			for (const cell of children(row, 'Cell')) {
				const cellName = attribute(cell, 'N');
				if (kind === 'User') {
					if (cellName === 'Value' && rowName) names.set(`User.${rowName}`, cell);
				} else if (kind === 'Geometry')
					names.set(`Geometry${index + 1}.${cellName}${attribute(row, 'IX')}`, cell);
				else {
					const prefix = kind === 'Connection' ? 'Connections' : 'Scratch';
					names.set(`${prefix}.${cellName}${Number(attribute(row, 'IX') ?? 0) + 1}`, cell);
				}
			}
		}
	}
	return names;
}

const lower = (name: string) => name.toLowerCase();
const byLowerName = (cells: Map<string, Element>) =>
	new Map([...cells].map(([name, cell]) => [lower(name), cell]));

/** Pair an instance with its master, recursively through the sub-shapes that name a master shape. */
function pairUp(inst: Element, master: Element, group: Group, out: Pair[]): void {
	const pair: Pair = {
		inst,
		master,
		cells: byLowerName(cellNames(inst)),
		masterCells: byLowerName(cellNames(master)),
		group,
	};
	out.push(pair);
	const masterId = attribute(inst, 'MasterShape');
	if (masterId) group.byMasterShape.set(masterId, pair);
	else if (attribute(inst, 'Master')) group.byMasterShape.set(attribute(master, 'ID') ?? '', pair);
	const subs = children(children(inst, 'Shapes')[0], 'Shape');
	const masterSubs = children(children(master, 'Shapes')[0], 'Shape');
	for (const sub of subs) {
		const target = masterSubs.find(
			(item) => attribute(item, 'ID') === attribute(sub, 'MasterShape'),
		);
		if (target) pairUp(sub, target, group, out);
	}
}

const isInherited = (cell: Element | undefined) => attribute(cell, 'F') === 'Inh';
const LENGTH =
	/^(width|height|pinx|piny|locpinx|locpiny|txtpinx|txtpiny|txtwidth|txtheight|txtlocpinx|txtlocpiny|geometry\d+\.[xy]\d+|connections\.[xy]\d+)$/;
/** A cached value; position and size cells are lengths whatever unit tag they carry. */
const numberOf = (cell: Element, name: string): VisioFormulaValue => {
	if (cell.hasAttribute('E')) throw new Error('error value');
	const value = visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'));
	return LENGTH.test(name.toLowerCase()) ? { value: value.value, unit: 'length' } : value;
};

class Recalc {
	private readonly values = new Map<string, VisioFormulaValue>();
	private readonly active = new Set<string>();
	/** Page shapes by id, for formulas the page itself carries (`Sheet.4!User.msvSDListDirection`). */
	private readonly pageShapes = new Map<string, Pair>();
	readonly pairs: Pair[] = [];
	constructor(
		private readonly root: Element,
		private readonly masters: CffMasters,
	) {
		const visit = (shapes: Element[]) => {
			for (const shape of shapes) {
				const masterId = attribute(shape, 'Master');
				const master = masterId ? this.masters.get(masterId) : undefined;
				if (master) {
					const group: Group = { byMasterShape: new Map() };
					const made: Pair[] = [];
					pairUp(shape, master, group, made);
					for (const pair of made) {
						this.pairs.push(pair);
						const id = attribute(pair.inst, 'ID');
						if (id) this.pageShapes.set(id, pair);
					}
				} else {
					const id = attribute(shape, 'ID');
					if (id)
						this.pageShapes.set(id, {
							inst: shape,
							master: shape,
							cells: byLowerName(cellNames(shape)),
							masterCells: new Map(),
							group: { byMasterShape: new Map() },
						});
					visit(children(children(shape, 'Shapes')[0], 'Shape'));
				}
			}
		};
		visit(children(children(root, 'Shapes')[0], 'Shape'));
	}

	/** The value of `name` on `pair`: its own cache, else the master's formula, else the master's cache. */
	value(pair: Pair, name: string, page: boolean): VisioFormulaValue {
		const key = `${attribute(pair.inst, 'ID')}!${lower(name)}`;
		const known = this.values.get(key);
		if (known) return known;
		if (this.active.has(key)) throw new Error('circular reference');
		this.active.add(key);
		try {
			const local = pair.cells.get(lower(name));
			const inherited = pair.masterCells.get(lower(name));
			let result: VisioFormulaValue | undefined;
			const formula = local && !isInherited(local) ? attribute(local, 'F') : undefined;
			if (formula) {
				try {
					result = this.run(pair, formula, true);
				} catch {
					// A page formula this interpreter cannot run (DEPENDSON): its saved cache stands.
					result = numberOf(local!, name);
				}
			} else if (local && !isInherited(local)) result = numberOf(local, name);
			else if (inherited) {
				const masterFormula = attribute(inherited, 'F');
				try {
					if (masterFormula && masterFormula !== 'Inh')
						result = this.run(pair, masterFormula, false);
				} catch (error) {
					// A formula this interpreter cannot run: the instance's cache, else the master's.
					if (!local && !inherited) throw error;
				}
				result ??= numberOf((local ?? inherited)!, name);
			} else if (local) result = numberOf(local, name);
			if (!result) throw new Error(`no cell ${name}`);
			this.values.set(key, result);
			return result;
		} finally {
			this.active.delete(key);
		}
	}

	private run(pair: Pair, formula: string, page: boolean): VisioFormulaValue {
		const out = evaluateVisioFormula(
			parseVisioFormula(formula),
			(ref: VisioFormulaReference) => {
				const lowered = ref.cell.toLowerCase();
				if (lowered === 'droponpagescale' || lowered === 'true')
					return { value: 1, unit: 'scalar' };
				if (lowered === 'false') return { value: 0, unit: 'scalar' };
				const target = ref.shapeId
					? page
						? this.pageShapes.get(ref.shapeId)
						: pair.group.byMasterShape.get(ref.shapeId)
					: pair;
				if (!target) throw new Error('unknown sheet');
				return this.value(target, ref.cell, page);
			},
			{ bareLengths: true },
		);
		return out;
	}

	/** Forget computed values (the caller changed cells). */
	reset(): void {
		this.values.clear();
	}
}

const SECTION_ORDER = [
	'FillGradient',
	'LineGradient',
	'Character',
	'Paragraph',
	'Tabs',
	'Scratch',
	'Connection',
	'Control',
	'User',
	'Property',
	'Hyperlink',
	'Geometry',
	'Actions',
	'Layer',
];
const rank = (section: Element) => {
	const index = SECTION_ORDER.indexOf(attribute(section, 'N') ?? '');
	return index < 0 ? SECTION_ORDER.length : index;
};
const create = (parent: Element, name: string) =>
	parent.ownerDocument.createElementNS(VISIO_NS, name);

/** The instance's copy of a master cell (`F="Inh"`), made with its section and row when missing. */
function ensureCell(inst: Element, master: Element): Element {
	const find = (parent: Element) =>
		children(parent, 'Cell').find((cell) => attribute(cell, 'N') === attribute(master, 'N'));
	const row = master.parentNode as Element;
	const top = row.localName === 'Shape';
	let holder: Element = inst;
	if (!top) {
		const masterSection = (row.localName === 'Section' ? row : row.parentNode) as Element;
		let section = children(inst, 'Section').find(
			(item) =>
				attribute(item, 'N') === attribute(masterSection, 'N') &&
				(attribute(item, 'IX') ?? '0') === (attribute(masterSection, 'IX') ?? '0'),
		);
		if (!section) {
			section = create(inst, 'Section');
			for (const name of ['N', 'IX'])
				if (masterSection.hasAttribute(name))
					section.setAttribute(name, attribute(masterSection, name)!);
			const after = children(inst, 'Section').find((item) => rank(item) > rank(section!));
			const anchor = after ?? children(inst, 'Text')[0] ?? children(inst, 'Shapes')[0] ?? null;
			inst.insertBefore(section, anchor);
		}
		holder = section;
		if (row.localName === 'Row') {
			let target = children(section, 'Row').find(
				(item) =>
					(attribute(item, 'IX') ?? '') === (attribute(row, 'IX') ?? '') &&
					(attribute(item, 'N') ?? '') === (attribute(row, 'N') ?? ''),
			);
			if (!target) {
				target = create(inst, 'Row');
				for (const name of ['T', 'N', 'IX'])
					if (row.hasAttribute(name)) target.setAttribute(name, attribute(row, name)!);
				section.appendChild(target);
			}
			holder = target;
		}
	}
	let cell = find(holder);
	if (!cell) {
		cell = create(inst, 'Cell');
		cell.setAttribute('N', attribute(master, 'N') ?? '');
		const unit = attribute(master, 'U');
		if (unit) cell.setAttribute('U', unit);
		cell.setAttribute('F', 'Inh');
		if (top) {
			const last = children(inst, 'Cell').pop();
			inst.insertBefore(cell, last ? last.nextSibling : inst.firstChild);
		} else holder.appendChild(cell);
	}
	return cell;
}

const same = (a: number, b: number) =>
	Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));

/**
 * Recompute the cached values of every master instance under `root` (a page's contents) from the
 * cells the instances carry. Returns the number of cells written. Cells whose formula cannot be
 * evaluated are left as they are.
 */
export function recalculateMasterInstances(
	root: Element,
	masters: CffMasters,
	only?: (instance: Element) => boolean,
): number {
	const engine = new Recalc(root, masters);
	let written = 0;
	for (const pair of engine.pairs) {
		if (only && !only(pair.inst)) continue;
		const names = new Set<string>();
		for (const [name, cell] of pair.masterCells)
			if (attribute(cell, 'F') && attribute(cell, 'F') !== 'Inh') names.add(name);
		for (const [name, cell] of pair.cells) if (attribute(cell, 'F')) names.add(name);
		for (const name of names) {
			let result: VisioFormulaValue;
			try {
				result = engine.value(pair, name, false);
			} catch (error) {
				continue;
			}
			const local = pair.cells.get(name);
			if (local) {
				if (!same(Number(attribute(local, 'V')), result.value)) {
					local.setAttribute('V', String(result.value));
					written++;
				}
				continue;
			}
			const master = pair.masterCells.get(name);
			const base = master && Number(attribute(master, 'V'));
			if (master && base !== undefined && !same(base, result.value)) {
				const cell = ensureCell(pair.inst, master);
				cell.setAttribute('V', String(result.value));
				written++;
			}
		}
	}
	return written;
}

/**
 * Give a top-level master instance local copies of the named cells (`Width`, `Height`, `LocPinX`
 * ...) holding their effective values, so code that reads an instance's own cells sees the
 * size and pin the master gives it. The cells are what Visio would save for the same shape.
 */
export function materializeCells(
	root: Element,
	masters: CffMasters,
	instance: Element,
	names: readonly string[],
): void {
	const engine = new Recalc(root, masters);
	const pair = engine.pairs.find((item) => item.inst === instance);
	if (!pair) return;
	for (const name of names) {
		const key = name.toLowerCase();
		if (pair.cells.has(key)) continue;
		const master = pair.masterCells.get(key);
		if (!master) continue;
		const result = engine.value(pair, key, false);
		ensureCell(instance, master).setAttribute('V', String(result.value));
	}
}
