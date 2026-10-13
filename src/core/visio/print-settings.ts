import type { VisioPageSetup, VisioPrintCountRefusal } from './model';
import { attribute, children } from './sheet';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula';

const NAMES = [
	'PaperKind',
	'PrintPageOrientation',
	'PageLeftMargin',
	'PageRightMargin',
	'PageTopMargin',
	'PageBottomMargin',
	'ScaleX',
	'ScaleY',
	'OnPage',
	'PagesX',
	'PagesY',
	'CenterX',
	'CenterY',
] as const;
const BOOLEAN = new Set(['OnPage', 'CenterX', 'CenterY']);
const LENGTH = new Set(['PageLeftMargin', 'PageRightMargin', 'PageTopMargin', 'PageBottomMargin']);
const CANONICAL = new Map(
	[...NAMES, 'PageWidth', 'PageHeight', 'PageScale', 'DrawingScale'].map((name) => [
		name.toLowerCase(),
		name,
	]),
);
interface IndexedCells {
	cells: Map<string, Element[]>;
	duplicate: boolean;
}
function indexCells(sheet: Element | undefined): IndexedCells {
	const cells = new Map<string, Element[]>();
	let duplicate = false;
	for (const node of children(sheet, 'Cell')) {
		const name = attribute(node, 'N') ?? '',
			canonical = CANONICAL.get(name.toLowerCase());
		if (!canonical) continue;
		if (canonical !== name) duplicate = true;
		const values = cells.get(canonical) ?? [];
		if (values.length < 2) values.push(node);
		cells.set(canonical, values);
		if (values.length > 1) duplicate = true;
	}
	return { cells, duplicate };
}

/** Index the document defaults once per parse, rather than scanning them for each page. */
export function createVisioPrintSettingsReader(
	documentSheets: readonly Element[],
	rootStyles: readonly Element[],
) {
	const indexes = { document: indexCells(documentSheets[0]), root: indexCells(rootStyles[0]) };
	return (pageSheets: readonly Element[]) =>
		readVisioPrintCountSettings(pageSheets, documentSheets, rootStyles, indexes);
}

/** Only count metadata inherits here. Shape/drawing resolution and writers do not use it. */
export function readVisioPrintCountSettings(
	pageSheets: readonly Element[],
	documentSheets: readonly Element[],
	rootStyles: readonly Element[],
	indexes?: { document: IndexedCells; root: IndexedCells },
): Partial<VisioPageSetup> {
	let refusal: VisioPrintCountRefusal | undefined;
	const refuse = (reason: VisioPrintCountRefusal) => {
		refusal ??= reason;
	};
	if (pageSheets.length !== 1) refuse('ambiguous-page-sheet');
	if (documentSheets.length > 1) refuse('ambiguous-document-sheet');
	if (!documentSheets.length) refuse('missing-setting');
	if (rootStyles.length > 1) refuse('duplicate-setting');
	const root = rootStyles[0];
	if (!root || attribute(root, 'ID') !== '0' || attribute(root, 'NameU') !== 'No Style')
		refuse('unsupported-print-style');
	for (const sheet of [...pageSheets, ...documentSheets])
		for (const name of ['LineStyle', 'FillStyle', 'TextStyle'])
			if (attribute(sheet, name) !== '0') refuse('unsupported-print-style');
	for (const name of ['LineStyle', 'FillStyle', 'TextStyle'])
		if (attribute(root, name) !== undefined) refuse('unsupported-print-style');
	const localIndex = indexCells(pageSheets[0]),
		rootIndex = indexes?.root ?? indexCells(root);
	const documentIndex = indexes?.document ?? indexCells(documentSheets[0]);
	if (localIndex.duplicate || rootIndex.duplicate || documentIndex.duplicate)
		refuse('duplicate-setting');
	const local = localIndex.cells,
		inherited = rootIndex.cells;
	// Native direct DocumentSheet print cells do not govern page settings. Do not
	// merge them as defaults; still retain the conservative duplicate-name guard.
	const read = (node: Element, name: string, origin: 'local' | 'document'): number | undefined => {
		const invalid = () => {
			refuse(origin === 'local' ? 'invalid-local-setting' : 'invalid-document-setting');
			return undefined;
		};
		if (node.hasAttribute('E')) return invalid();
		let raw = (attribute(node, 'V') ?? '').trim();
		if (!raw || raw.length > 128) return invalid();
		if (BOOLEAN.has(name)) {
			if (!['0', '1', 'false', 'true'].includes(raw)) return invalid();
			raw = raw === 'true' ? '1' : raw === 'false' ? '0' : raw;
		}
		try {
			const value = visioFormulaCachedValue(raw, attribute(node, 'U'));
			const length =
				LENGTH.has(name) || /^(PageWidth|PageHeight|PageScale|DrawingScale)$/.test(name);
			if (
				!['scalar', ...(length ? ['length'] : [])].includes(value.unit) ||
				Math.abs(value.value) > 1e9
			)
				return invalid();
			const formula = attribute(node, 'F');
			if (formula !== undefined) {
				const source = formula.trim();
				if (/^(true|false)$/i.test(source)) {
					if ((source.toLowerCase() === 'true' ? 1 : 0) !== value.value) return invalid();
				} else {
					if (!source || source === 'Inh') return invalid();
					const limits = { maxLength: 512, maxNodes: 64, maxDepth: 16, maxSteps: 128 };
					const analysis = analyzeVisioFormula(source, limits);
					if (
						analysis.dynamic ||
						analysis.guarded ||
						analysis.references.length ||
						analysis.unsupportedFunctions.length
					) {
						refuse('dynamic-setting');
						return undefined;
					}
					const evaluated = evaluateVisioFormula(
						source,
						() => {
							throw new Error('Unexpected reference');
						},
						limits,
					);
					if (
						evaluated.value !== value.value ||
						!['scalar', ...(length ? ['length'] : [])].includes(evaluated.unit)
					)
						return invalid();
				}
			}
			return value.value;
		} catch {
			return invalid();
		}
	};
	const values = new Map<string, number>();
	for (const name of NAMES) {
		const nodes = local.get(name) ?? inherited.get(name);
		if (!nodes?.length) continue;
		// An explicit unusable local cell must never fall back to No Style defaults.
		if (nodes.length !== 1) continue;
		const value = read(nodes[0]!, name, local.has(name) ? 'local' : 'document');
		if (value !== undefined) values.set(name, value);
	}
	// Physical extents already use the existing page geometry projection. Prove its inputs;
	// do not import document drawing defaults or alter the shape-resolution path.
	for (const name of ['PageWidth', 'PageHeight', 'PageScale', 'DrawingScale']) {
		const nodes = local.get(name);
		if (!nodes || nodes.length !== 1 || !(read(nodes[0]!, name, 'local')! > 0))
			refuse('unusable-drawing-scale');
	}
	const setup: { -readonly [K in keyof VisioPageSetup]: VisioPageSetup[K] } = {};
	const paper = values.get('PaperKind');
	if (paper !== undefined && Number.isSafeInteger(paper) && paper >= 0 && paper <= 0xffff)
		setup.paperKind = paper;
	else if (paper !== undefined)
		refuse(local.has('PaperKind') ? 'invalid-local-setting' : 'invalid-document-setting');
	const orientation = values.get('PrintPageOrientation');
	if (orientation === 0 || orientation === 1 || orientation === 2)
		setup.printPageOrientation = orientation;
	else if (orientation !== undefined)
		refuse(
			local.has('PrintPageOrientation') ? 'invalid-local-setting' : 'invalid-document-setting',
		);
	const margins = ['Left', 'Right', 'Top', 'Bottom'].map((side) => values.get(`Page${side}Margin`));
	if (margins.every((n) => n !== undefined && n >= 0 && n <= 1000))
		setup.margins = {
			left: margins[0]!,
			right: margins[1]!,
			top: margins[2]!,
			bottom: margins[3]!,
		};
	else if (margins.some((n) => n !== undefined && (n < 0 || n > 1000)))
		refuse('unusable-print-area');
	for (const [name, key] of [
		['ScaleX', 'printZoom'],
		['ScaleY', 'printZoomY'],
		['PagesX', 'printPagesAcross'],
		['PagesY', 'printPagesDown'],
	] as const) {
		const value = values.get(name);
		if (value !== undefined) setup[key] = value;
	}
	for (const [name, key] of [
		['OnPage', 'printFitToPages'],
		['CenterX', 'printCenterX'],
		['CenterY', 'printCenterY'],
	] as const) {
		const value = values.get(name);
		if (value !== undefined) setup[key] = value === 1;
	}
	if (refusal) setup.printCountRefusal = refusal;
	return setup;
}
