import { textSizeFormula } from './edit-text-size-formula';
import { isLocalForeignFrameShape } from './edit-foreign-frame';
import { children, attribute } from './sheet';
import { fail } from './package-common';
import { independentTextFieldFormula } from './text-field-independent';
import { parseVisioFormula, analyzeVisioFormula, type VisioFormulaValue } from './formula';
import {
	executableCellFormula,
	inertDoubleClickFormula,
	shapeDataFormulaSection,
} from './cell-formula';
import { isClassifiedVisioFunction } from './formula-analyze';
import { isNativeGlueCell } from './edit-connector-glue';

export interface VisioCellKey {
	pageId: string;
	shapeId: string;
	cell: string;
}
export interface VisioRecalculationOptions {
	pageContext?: ReadonlyMap<string, { pageNumber: number; pageCount: number }>;
	maxCells?: number;
	maxAffectedCells?: number;
	maxSteps?: number;
	maxDepth?: number;
	check?: () => void;
	/** Internal post-proof authorization of explicit local master move leaves only. */
	masterMovePins?: ReadonlySet<Element>;
	/** Internal authorization after proving an unglued local straight-line edit. */
	lineEditShapes?: ReadonlySet<Element>;
	/** Internal authorization of proven owned dynamic-glue connectors and their 2D targets. */
	glueShapes?: ReadonlySet<Element>;
	/** Internal authorization of the proven local parent-group Angle leaf only. */
	groupRotationCells?: ReadonlySet<Element>;
	/** Internal group-resize scope: read-only dependency discovery, or a proven local writer. */
	groupResizeShapes?: ReadonlySet<Element>;
	/**
	 * A cell caught in a dependency cycle keeps its cached value and is reported in `diagnostics`
	 * instead of refusing the whole recalculation. Default: refuse.
	 */
	cycles?: 'refuse' | 'keep';
	/**
	 * `affected`: a cell whose formula cannot be analysed (dynamic functions, unknown functions)
	 * is indexed as unevaluable instead of failing the whole index, so it refuses an edit only when
	 * the edit reaches it. A cell that may read cells its text does not name is assumed to be
	 * reached by every edit on its page. Default `document`: such a cell fails the index.
	 */
	scope?: 'document' | 'affected';
	/**
	 * Follow what the evaluator cannot prove statically instead of refusing: a computed INDIRECT or
	 * EVALCELL is resolved while evaluating, a function it cannot run keeps the cached result (and is
	 * reported in `diagnostics`), and a cell of a grouped, instanced or connected shape that the edit
	 * does not reach is read as an input. Set by `recalculateVisioCells`; other callers keep refusing.
	 */
	follow?: boolean;
	/** Receives one entry per cell that kept its cache because of a cycle. */
	diagnostics?: { code: string; message: string }[];
	/** The plain text of a shape for `TheText` readers; recalculation fills it from the pages. */
	textOf?: (pageId: string, shapeId: string) => string | undefined;
	/** `follow` only: the cells the edit reaches; any other cell of an unsafe shape is an input. */
	affected?: ReadonlySet<string>;
}
export interface IndexedCell extends VisioCellKey {
	node?: Element;
	formula?: ReturnType<typeof parseVisioFormula>;
	dependencies: string[];
	unsafe: boolean;
	/** Why the cell cannot be evaluated, when its formula was indexed as unevaluable. */
	refusal?: string;
	/** The formula may read cells it does not name: any edit on its page may reach it. */
	opaque?: boolean;
	/** `follow` only: the formula cannot be run here; its cached result stands in. */
	keepCache?: boolean;
	/** `follow` only: the cell sits on an unsafe shape but is a plain input when no edit reaches it. */
	fixedInput?: boolean;
	unit: VisioFormulaValue['unit'];
}
export const key = (cell: VisioCellKey) =>
	JSON.stringify([cell.pageId, cell.shapeId, cell.cell.toLowerCase()]);
export function bound(value: number | undefined, fallback: number): number {
	const result = value ?? fallback;
	if (!Number.isSafeInteger(result) || result < 1)
		fail('INVALID_LIMITS', 'Invalid recalculation limit.');
	return result;
}
function cellUnit(name: string, relative: boolean): VisioFormulaValue['unit'] {
	if (/^(Angle|TxtAngle)$/i.test(name)) return 'angle';
	if (
		/^(Width|Height|PinX|PinY|LocPinX|LocPinY|BeginX|BeginY|EndX|EndY|TxtPinX|TxtPinY|TxtWidth|TxtHeight|TxtLocPinX|TxtLocPinY|LineWeight|Rounding|ImgOffsetX|ImgOffsetY|ImgWidth|ImgHeight)$/i.test(
			name,
		)
	)
		return 'length';
	return !relative && /\.[XY]\d+$/i.test(name) ? 'length' : 'scalar';
}
/** All local sheets participate, including sections, named rows, and static Sheet.ID references. */
/** `TRUE` and `FALSE` parse as names; they are the ShapeSheet's constants 1 and 0, not cells. */
export const isBooleanLiteral = (ref: { shapeId?: string; cell: string }): boolean =>
	ref.shapeId === undefined && /^(true|false)$/i.test(ref.cell);

export function indexCells(
	roots: ReadonlyMap<string, Element>,
	options: VisioRecalculationOptions,
) {
	const cells = new Map<string, IndexedCell>();
	const maxCells = bound(options.maxCells, 100_000);
	const add = (item: IndexedCell) => {
		const id = key(item);
		if (cells.has(id)) fail('EDIT_AMBIGUOUS_CELL', 'Duplicate local ShapeSheet cell.');
		cells.set(id, item);
		if (cells.size > maxCells) fail('LIMIT_FORMULA_CELLS', 'ShapeSheet cell limit exceeded.');
	};
	for (const [pageId, root] of roots) {
		if (root.hasAttribute('F'))
			fail('EDIT_UNKNOWN_DEPENDENCY', 'Page root formula has unknown dependency scope.');
		const indexedNodes = new Set<Element>();
		const connected = new Set<string>(),
			connectors = new Set<string>();
		for (const container of children(root, 'Connects'))
			for (const connection of children(container, 'Connect')) {
				for (const attr of ['FromSheet', 'ToSheet']) {
					const id = attribute(connection, attr);
					if (id) connected.add(id);
				}
				const from = attribute(connection, 'FromSheet');
				if (from) connectors.add(from);
			}
		const sheets: { node: Element; shapeId: string; unsafe: boolean }[] = [];
		const ids = new Set<string>();
		const visit = (parent: Element, inherited: boolean, nested: boolean) => {
			for (const container of children(parent, 'Shapes'))
				for (const node of children(container, 'Shape')) {
					options.check?.();
					const shapeId = attribute(node, 'ID');
					if (!shapeId || ids.has(shapeId)) fail('INVALID_SHAPE_ID', 'Shape IDs must be unique.');
					ids.add(shapeId);
					const unsafe =
						inherited ||
						nested ||
						node.hasAttribute('Master') ||
						node.hasAttribute('MasterShape') ||
						attribute(node, 'Type') === 'Group' ||
						(!isLocalForeignFrameShape(node) &&
							((attribute(node, 'Type') !== undefined && attribute(node, 'Type') !== 'Shape') ||
								children(node, 'ForeignData').length > 0 ||
								children(node, 'Rel').length > 0)) ||
						(connected.has(shapeId) && !options.glueShapes?.has(node)) ||
						(!options.lineEditShapes?.has(node) &&
							children(node, 'Cell').some(
								(cell) =>
									/^(BeginX|BeginY|EndX|EndY)$/i.test(attribute(cell, 'N') ?? '') ||
									(attribute(cell, 'N')?.toLowerCase() === 'oned' &&
										(attribute(cell, 'V') !== '0' ||
											!!executableCellFormula(attribute(cell, 'F')) ||
											attribute(cell, 'F') === 'Inh')),
							)) ||
						children(node, 'Shapes').length > 0 ||
						['1', 'true'].includes(attribute(node, 'Del') ?? '');
					sheets.push({
						node,
						shapeId,
						unsafe: options.groupResizeShapes?.has(node) ? false : unsafe,
					});
					visit(node, unsafe, true);
				}
		};
		visit(root, false, false);
		for (const node of children(root, 'PageSheet'))
			sheets.push({ node, shapeId: '', unsafe: false });
		for (const sheet of sheets) {
			const addNode = (
				node: Element,
				cell: string,
				relative = false,
				deleted = false,
				glue = false,
			) => {
				indexedNodes.add(node);
				const item: IndexedCell = {
					pageId,
					shapeId: sheet.shapeId,
					cell,
					node,
					unsafe: sheet.unsafe || deleted,
					dependencies: [],
					unit: cellUnit(cell, relative),
				};
				const source = attribute(node, 'F');
				// Native dynamic glue is re-evaluated by Visio; the editor reads and writes its caches.
				if (
					executableCellFormula(source) &&
					!independentTextFieldFormula(node, source!) &&
					!inertDoubleClickFormula(cell, source!) &&
					!(glue && isNativeGlueCell(cell, source!)) &&
					// A trigger only counts events of the sheet it names; an unglued connector keeps one
					// (Visio's freshly dropped connector names its own sheet).
					!(/^(BegTrigger|EndTrigger)$/.test(cell) && isNativeGlueCell(cell, source!))
				) {
					const lenient = options.scope === 'affected';
					const unprovable = (message: string, opaque: boolean) => {
						if (!lenient) fail('EDIT_UNKNOWN_DEPENDENCY', message);
						item.unsafe = true;
						item.refusal = message;
						if (opaque) item.opaque = true;
					};
					try {
						item.formula = parseVisioFormula(source!);
						// A size measured from the shape's own text reads only the cells it names; the
						// text itself is followed after text edits (edit-text-size.ts).
						const plain = analyzeVisioFormula(item.formula, {
							shapeTextReads: true,
							literalIndirect: true,
						});
						const analysis = (plain.dynamic && textSizeFormula(item.formula)) || plain;
						const follows = options.follow === true && lenient;
						const dynamicCalls = 'dynamicCalls' in analysis ? analysis.dynamicCalls : [];
						const resolvable = dynamicCalls.every(
							(name) => name === 'INDIRECT' || name === 'EVALCELL',
						);
						if (follows && analysis.unsupportedFunctions.length === 0 && resolvable) {
							// Resolved per evaluation; the reads the text does not name make it page-wide.
							if (analysis.dynamic) item.opaque = true;
						} else if (
							follows &&
							(analysis.dynamic || analysis.unsupportedFunctions.length) &&
							!textSizeFormula(item.formula) &&
							// Drawn geometry data: a stale copy would draw the wrong outline.
							!analysis.unsupportedFunctions.some((name) => name === 'POLYLINE' || name === 'NURBS')
						) {
							// DEPENDSON names its dependencies; the other dynamic functions may read more.
							item.keepCache = true;
							if (
								dynamicCalls.some((name) => name !== 'DEPENDSON' && isClassifiedVisioFunction(name))
							)
								item.opaque = true;
						} else if (analysis.dynamic) {
							if (!lenient)
								fail(
									'EDIT_DYNAMIC_DEPENDENCY',
									'Dynamic ShapeSheet dependencies cannot be safely edited.',
								);
							unprovable(
								`${cell} on shape ${sheet.shapeId || 'page'} has dynamic ShapeSheet dependencies that cannot be followed.`,
								'opaque' in analysis && analysis.opaque,
							);
						}
						item.dependencies = analysis.references
							.filter((ref) => !isBooleanLiteral(ref))
							.map((ref) => key({ pageId, shapeId: ref.shapeId ?? sheet.shapeId, cell: ref.cell }));
					} catch (error) {
						unprovable(
							`Cannot prove ShapeSheet dependency closure: ${error instanceof Error ? error.message : 'invalid formula'}`,
							true,
						);
					}
				}
				if (source === 'Inh') item.unsafe = true;
				else if (options.follow && sheet.unsafe && !deleted && !item.refusal)
					item.fixedInput = true;
				add(item);
			};
			const glue = connectors.has(sheet.shapeId);
			for (const node of children(sheet.node, 'Cell'))
				addNode(node, attribute(node, 'N') ?? '', false, false, glue);
			for (const section of children(sheet.node, 'Section')) {
				// VSDX stores Shape Data as Property; ShapeSheet formulas address it as Prop.
				const name = shapeDataFormulaSection(attribute(section, 'N') ?? '');
				const prefix =
					name === 'Geometry' ? `Geometry${Number(attribute(section, 'IX') ?? '0') + 1}` : name;
				const deleted = ['1', 'true'].includes(attribute(section, 'Del') ?? '');
				for (const node of children(section, 'Cell'))
					addNode(node, `${prefix}.${attribute(node, 'N') ?? ''}`, false, deleted);
				for (const row of children(section, 'Row')) {
					const rowName = attribute(row, 'N');
					const ix = attribute(row, 'IX') ?? '0';
					for (const node of children(row, 'Cell')) {
						const n = attribute(node, 'N') ?? '';
						const rowCell =
							name === 'Geometry' || ['Controls', 'Scratch', 'Connection'].includes(name)
								? `${prefix}.${n}${ix}`
								: n === 'Value' && ['User', 'Prop'].includes(name)
									? `${prefix}.${rowName ?? ix}`
									: `${prefix}.${rowName ?? ix}.${n}`;
						addNode(
							node,
							rowCell,
							(attribute(row, 'T') ?? '').startsWith('Rel'),
							deleted || ['1', 'true'].includes(attribute(row, 'Del') ?? ''),
						);
					}
				}
			}
			if (sheet.shapeId)
				for (const [cell, dimension, scale] of [
					['LocPinX', 'Width', 0.5],
					['LocPinY', 'Height', 0.5],
					['TxtWidth', 'Width', 1],
					['TxtHeight', 'Height', 1],
					['TxtPinX', 'Width', 0.5],
					['TxtPinY', 'Height', 0.5],
					['TxtLocPinX', 'TxtWidth', 0.5],
					['TxtLocPinY', 'TxtHeight', 0.5],
				] as const) {
					const ref = { pageId, shapeId: sheet.shapeId, cell };
					if (!cells.has(key(ref)))
						add({
							...ref,
							unsafe: sheet.unsafe,
							unit: 'length',
							dependencies: [key({ ...ref, cell: dimension })],
							formula: parseVisioFormula(`${dimension}*${scale}`),
						});
				}
		}
		// Formula-bearing markup outside admitted sheet cells has no proven dependency scope.
		for (const node of Array.from(root.getElementsByTagName('*'))) {
			options.check?.();
			if (node.hasAttribute('F') && !indexedNodes.has(node))
				fail(
					'EDIT_UNKNOWN_DEPENDENCY',
					'Formula outside admitted ShapeSheet cells has unknown dependencies.',
				);
		}
	}
	return cells;
}
