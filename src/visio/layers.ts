import type { VisioLayer, VisioLayerPrintSummary } from './model.js';
import { metadata } from './metadata.js';
import { VisioPackageError } from './package.js';
import { number, sectionRows, type Report, type Sheet } from './sheet.js';
import { color, type Resources } from './style.js';

export interface LayerBudget {
	consumeCharacters(count: number): void;
	consumeMemberships(count: number): void;
}
/** Share across pages and inherited instances to bound normalization work and arrays. */
export function createLayerBudget(): LayerBudget {
	let memberships = 0,
		characters = 0;
	return {
		consumeCharacters(count) {
			characters += count;
			if (characters > 5_000_000)
				throw new VisioPackageError(
					'METADATA_LIMIT',
					'Layer membership character budget exceeded.',
				);
		},
		consumeMemberships(count) {
			memberships += count;
			if (memberships > 100_000)
				throw new VisioPackageError('METADATA_LIMIT', 'Layer membership budget exceeded.');
		},
	};
}
/** Construct once per page; preserve the first matching row's existing semantics. */
export function indexLayers(layers: readonly VisioLayer[]): ReadonlyMap<string, VisioLayer> {
	const index = new Map<string, VisioLayer>();
	for (const layer of layers) if (!index.has(layer.id)) index.set(layer.id, layer);
	return index;
}
/** PageSheet Layer row indices are the stable layer IDs used by LayerMember. */
export function pageLayers(sheet: Sheet, resources: Resources, report: Report): VisioLayer[] {
	return sectionRows(sheet, 'Layer').map((row) => {
		const colorValue = row.cells.get('Color')?.value;
		return {
			id: row.index,
			name: metadata(
				row.cells.get('Name')?.value ?? row.cells.get('NameUniv')?.value ?? `Layer ${row.index}`,
				4096,
				'Layer name',
			),
			visible: number(row.cells, 'Visible', 1, report) !== 0,
			printable: number(row.cells, 'Print', 1, report) !== 0,
			locked: number(row.cells, 'Lock', 0, report) !== 0,
			...(colorValue === undefined || colorValue === '255'
				? {}
				: {
						color: color(row.cells, 'Color', '#000000', resources, report),
						colorOpacity: Math.max(0, Math.min(1, 1 - number(row.cells, 'ColorTrans', 0, report))),
					}),
		};
	});
}
export function shapeLayers(
	sheet: Sheet,
	layers: ReadonlyMap<string, VisioLayer>,
	report: Report,
	budget: LayerBudget = createLayerBudget(),
): { layerIds: string[]; hidden: boolean; printSummary: VisioLayerPrintSummary } {
	const cached = sheet.cells.get('LayerMember');
	const raw = metadata(cached?.value ?? '', 8192, 'Cached layer membership');
	budget.consumeCharacters(raw.length);
	const value = raw.trim();
	if (!value)
		return {
			layerIds: [],
			hidden: false,
			printSummary: cached && cached.value === undefined ? 'unknown' : 'unlayered',
		};
	let count = 1;
	for (const character of value)
		if (character === ';' && ++count > 1024)
			throw new VisioPackageError('METADATA_LIMIT', 'Shape layer membership count exceeded.');
	budget.consumeMemberships(count);
	if (!/^\d+(;\d+)*$/.test(value)) {
		report('invalid-layer-membership', 'Invalid cached layer membership was ignored.');
		return { layerIds: [], hidden: false, printSummary: 'unknown' };
	}
	// LayerMember permits decimal digits. Preserve unknown large IDs without Number rounding.
	const layerIds = [
		...new Set(
			value.split(';').map((id) => metadata(id.replace(/^0+/, '') || '0', 256, 'Layer member ID')),
		),
	];
	let missing = false,
		colored = false,
		hidden = false,
		printEnabled = false,
		printDisabled = false;
	for (const id of layerIds) {
		const layer = layers.get(id);
		missing ||= layer === undefined;
		colored ||= layer?.color !== undefined;
		hidden ||= layer?.visible === false;
		printEnabled ||= layer?.printable === true;
		printDisabled ||= layer?.printable === false;
	}
	if (missing) report('missing-layer', 'Shape references a layer absent from this page.');
	if (colored)
		report(
			'unsupported-layer-color',
			'Layer color overrides are retained as metadata but not applied to shape styling.',
		);
	// MS-VSDX 2.2.3.2.2: a geometry path must not belong to a layer whose Visible is zero.
	return {
		layerIds,
		hidden,
		// Summarize normalized Print flags, including pageLayers' existing cached-value defaults.
		// This deliberately does not resolve mixed print policy or a shape's final printability.
		printSummary: missing
			? 'unknown'
			: printEnabled && printDisabled
				? 'mixed'
				: printEnabled
					? 'all-enabled'
					: 'all-disabled',
	};
}
