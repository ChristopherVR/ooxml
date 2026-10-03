import {
	PROPERTY_CONTENT_TYPES,
	PROPERTY_PART_NAMES,
	RELATIONSHIP_TYPES,
	writeAppProperties,
	writeCoreProperties,
	writeCustomProperties,
	type AppProperties,
	type HeadingPair,
} from '../../opc/index.js';
import type { Workbook } from '../model.js';
import { CONTENT_TYPES, SourceIndex } from '../read/package.js';
import type { PackageWriter, RelationshipSet } from './package-writer.js';

/** `HeadingPairs` groups whose titles are sheet names (Excel writes these English names). */
const SHEET_GROUPS = new Set(['worksheets', 'charts', 'dialogs', 'macros', 'excel 4.0 macros']);

const indexOf = (workbook: Workbook, source?: SourceIndex): SourceIndex | undefined =>
	source ?? (workbook.source ? new SourceIndex(workbook.source.parts) : undefined);

function sourceText(source: SourceIndex | undefined, type: string): string | undefined {
	const name = source?.targetOfType('', type);
	return name ? source?.text(name) : undefined;
}

/**
 * `HeadingPairs` and `TitlesOfParts` for the sheets being saved: the model's groups are kept
 * when their sheet titles still match the workbook's sheets, otherwise the sheet groups are
 * replaced by one `Worksheets` group. Non-sheet groups (`Named Ranges`) are kept as they are.
 */
export function sheetTitles(workbook: Workbook): { pairs: HeadingPair[]; titles: string[] } {
	const names = workbook.sheets.map((sheet) => sheet.name);
	const pairs = workbook.properties.headingPairs ?? [];
	const titles = workbook.properties.titlesOfParts ?? [];
	const groups: { pair: HeadingPair; titles: string[] }[] = [];
	let offset = 0;
	for (const pair of pairs) {
		groups.push({ pair, titles: titles.slice(offset, offset + pair.count) });
		offset += pair.count;
	}
	const consistent = offset === titles.length;
	const isSheet = (pair: HeadingPair) => SHEET_GROUPS.has(pair.name.toLowerCase());
	const sheetTitlesNow = groups.filter((g) => isSheet(g.pair)).flatMap((g) => g.titles);
	const unchanged =
		consistent &&
		sheetTitlesNow.length === names.length &&
		[...sheetTitlesNow].sort().join('\u0000') === [...names].sort().join('\u0000');
	if (unchanged) return { pairs, titles };
	const kept = consistent ? groups : [];
	const at = Math.max(
		0,
		kept.findIndex((g) => isSheet(g.pair)),
	);
	const rest = kept.filter((g) => !isSheet(g.pair));
	rest.splice(Math.min(at, rest.length), 0, {
		pair: { name: 'Worksheets', count: names.length },
		titles: names,
	});
	return { pairs: rest.map((g) => g.pair), titles: rest.flatMap((g) => g.titles) };
}

/** `docProps/core.xml`: the source part patched with the workbook properties. */
export function coreXml(workbook: Workbook, source?: SourceIndex): string {
	return writeCoreProperties(
		workbook.properties,
		sourceText(indexOf(workbook, source), RELATIONSHIP_TYPES.coreProperties),
	);
}

/**
 * `docProps/app.xml`: the source part patched with the workbook properties (elements the model
 * does not know are kept), sheet titles refreshed. A new workbook gets Excel's defaults.
 */
export function appXml(workbook: Workbook, sourceIndex?: SourceIndex): string {
	const source = sourceText(indexOf(workbook, sourceIndex), RELATIONSHIP_TYPES.extendedProperties);
	const p = workbook.properties;
	const { pairs, titles } = sheetTitles(workbook);
	const defaults: AppProperties = source
		? {}
		: {
				docSecurity: 0,
				scaleCrop: false,
				linksUpToDate: false,
				sharedDoc: false,
				hyperlinksChanged: false,
				appVersion: '16.0300',
			};
	const props: AppProperties = {
		...defaults,
		...Object.fromEntries(Object.entries(p).filter(([, value]) => value !== undefined)),
		application: p.application ?? 'Microsoft Excel',
		headingPairs: pairs,
		titlesOfParts: titles,
	};
	return writeAppProperties(props, source);
}

/** Whether the package-root relationship of `type` is rewritten by {@link writeDocProps}. */
export const regeneratesRootRel = (workbook: Workbook, type: string): boolean =>
	type === RELATIONSHIP_TYPES.customProperties && workbook.properties.custom !== undefined;

/**
 * Writes `docProps/core.xml` and `docProps/app.xml` (patched from the source parts) and, when
 * the model holds custom properties (`properties.custom` set), `docProps/custom.xml`, adding
 * their package-root relationships to `root`. With `properties.custom` absent the source custom
 * part, if any, is carried unchanged with the other root relationships.
 */
export function writeDocProps(
	writer: PackageWriter,
	root: RelationshipSet,
	workbook: Workbook,
	source: SourceIndex | undefined,
): void {
	writer.add(PROPERTY_PART_NAMES.core, coreXml(workbook, source), CONTENT_TYPES.core);
	writer.add(PROPERTY_PART_NAMES.app, appXml(workbook, source), CONTENT_TYPES.app);
	const custom = workbook.properties.custom;
	if (custom === undefined) return;
	const xml = writeCustomProperties(custom);
	if (xml === undefined) return;
	const name =
		source?.targetOfType('', RELATIONSHIP_TYPES.customProperties) ?? PROPERTY_PART_NAMES.custom;
	root.add(RELATIONSHIP_TYPES.customProperties, name);
	writer.add(name, xml, PROPERTY_CONTENT_TYPES.custom);
}
