import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import { DIGITAL_SIGNATURE_ORIGIN_REL_TYPE, isSignaturePart } from '../../opc/signature/index.js';
import type { Workbook } from '../model.js';
import { parsePersons } from '../read/comments.js';
import { CONTENT_TYPES, SourceIndex } from '../read/package.js';
import { PersonRegistry } from './comments.js';
import { relativeTarget } from './drawing.js';
import { patchCarriedParts, workbookRefEdits } from './carried-refs.js';
import { MetadataPlan, writeMetadataPart } from './metadata.js';
import { PackageWriter, RelationshipSet } from './package-writer.js';
import { SharedStringTable } from './shared-strings.js';
import { StyleWriter } from './styles.js';
import { themePart } from './theme.js';
import { regeneratesRootRel, writeDocProps } from './doc-props.js';
import { workbookXml, type SheetRef } from './workbook-part.js';
import { writeWorksheet, type SaveContext } from './worksheet.js';

/** Workbook relationships regenerated on save (everything else in the source is kept). */
const REGENERATED = new Set<string>([
	RELATIONSHIP_TYPES.worksheet,
	RELATIONSHIP_TYPES.chartsheet,
	RELATIONSHIP_TYPES.dialogsheet,
	RELATIONSHIP_TYPES.xlMacrosheet,
	RELATIONSHIP_TYPES.styles,
	RELATIONSHIP_TYPES.sharedStrings,
	RELATIONSHIP_TYPES.theme,
	RELATIONSHIP_TYPES.calcChain,
	RELATIONSHIP_TYPES.person,
	RELATIONSHIP_TYPES.sheetMetadata,
]);

const SHEET_TYPES: Record<string, string> = {
	[CONTENT_TYPES.chartsheet]: RELATIONSHIP_TYPES.chartsheet,
	'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml':
		RELATIONSHIP_TYPES.dialogsheet,
	'application/vnd.ms-excel.macrosheet+xml': RELATIONSHIP_TYPES.xlMacrosheet,
};

const ROOT_REGENERATED = new Set<string>([
	RELATIONSHIP_TYPES.officeDocument,
	RELATIONSHIP_TYPES.coreProperties,
	RELATIONSHIP_TYPES.extendedProperties,
]);

/**
 * Writes the workbook as an `.xlsx` (or `.xlsm` when it was loaded with VBA). Workbook, sheet,
 * style, string, table, comment and property parts are regenerated from the model; every
 * other part of the source package that is still referenced is copied unchanged.
 */
export async function saveXlsx(workbook: Workbook): Promise<Uint8Array> {
	if (!workbook.sheets.length) throw new Error('A workbook needs at least one sheet');
	const source = workbook.source ? new SourceIndex(workbook.source.parts) : undefined;
	const writer = new PackageWriter(source);
	const sourceBook = source?.workbookPart();
	const bookPart = sourceBook ?? 'xl/workbook.xml';
	const reserved = new Set(source?.parts.keys() ?? []);
	const sourcePart = (type: string) =>
		source && sourceBook ? source.targetOfType(sourceBook, type) : undefined;
	const stylesPart = sourcePart(RELATIONSHIP_TYPES.styles);
	const personsSource = sourcePart(RELATIONSHIP_TYPES.person);
	const metadataSource = sourcePart(RELATIONSHIP_TYPES.sheetMetadata);
	const metadata = new MetadataPlan(metadataSource ? source?.text(metadataSource) : undefined);
	const ctx: SaveContext = {
		writer,
		workbook,
		source,
		styles: new StyleWriter(workbook, stylesPart ? source?.text(stylesPart) : undefined),
		strings: new SharedStringTable(),
		persons: new PersonRegistry(
			parsePersons(personsSource ? source?.text(personsSource) : undefined),
		),
		reserved,
		tableIds: new Set(),
		threaded: { used: false },
		dynamicArrays: { used: false },
		metadata,
	};
	const rels = new RelationshipSet();
	let keptVba = false;
	if (source && sourceBook) {
		for (const [id, rel] of source.rels(sourceBook)) {
			if (REGENERATED.has(rel.type)) continue;
			if (rel.type === RELATIONSHIP_TYPES.vbaProject) {
				if (workbook.format !== 'xlsm') continue;
				keptVba = true;
			}
			rels.keep(id, rel);
			const target = source.target(sourceBook, rel);
			if (target) writer.carry(target);
		}
	}

	// Excel refuses a workbook whose sheets are all hidden.
	const states = workbook.sheets.map((sheet) => sheet.state);
	if (!states.includes('visible')) states[0] = 'visible';
	let activeTab = Math.min(Math.max(0, workbook.activeSheet), workbook.sheets.length - 1);
	if (states[activeTab] !== 'visible') activeTab = states.indexOf('visible');

	const claimed = new Set<string>();
	const partNames = workbook.sheets.map((sheet) => {
		const own = sheet.partName;
		if (
			own &&
			!claimed.has(own) &&
			(!source ||
				!source.has(own) ||
				source.contentType(own) === CONTENT_TYPES.worksheet ||
				SHEET_TYPES[source.contentType(own) ?? ''])
		) {
			claimed.add(own);
			return own;
		}
		return undefined;
	});
	const sheetRefs: SheetRef[] = [];
	workbook.sheets.forEach((sheet, index) => {
		let partName = partNames[index];
		const sourceType = partName && source?.has(partName) ? source.contentType(partName) : undefined;
		const special = sourceType ? SHEET_TYPES[sourceType] : undefined;
		if (!partName) {
			partName = writer.uniqueName(
				(n) => `xl/worksheets/sheet${n}.xml`,
				new Set([...reserved, ...claimed]),
			);
			claimed.add(partName);
		}
		if (special) writer.carry(partName);
		else writeWorksheet(ctx, sheet, index, partName);
		sheetRefs.push({
			name: sheet.name,
			sheetId: sheet.sheetId,
			state: states[index] ?? 'visible',
			relId: rels.add(special ?? RELATIONSHIP_TYPES.worksheet, relativeTarget(bookPart, partName)),
		});
	});

	const stringsPart = sourcePart(RELATIONSHIP_TYPES.sharedStrings) ?? 'xl/sharedStrings.xml';
	if (ctx.strings.size) {
		writer.add(stringsPart, ctx.strings.xml(), CONTENT_TYPES.sharedStrings);
		rels.add(RELATIONSHIP_TYPES.sharedStrings, relativeTarget(bookPart, stringsPart));
	}
	const outStyles = stylesPart ?? 'xl/styles.xml';
	writer.add(outStyles, ctx.styles.xml(), CONTENT_TYPES.styles);
	rels.add(RELATIONSHIP_TYPES.styles, relativeTarget(bookPart, outStyles));
	const sourceTheme = sourcePart(RELATIONSHIP_TYPES.theme);
	const theme = themePart(workbook.theme, sourceTheme ? source?.text(sourceTheme) : undefined);
	const outTheme = sourceTheme ?? 'xl/theme/theme1.xml';
	if (theme === undefined && sourceTheme) writer.carry(sourceTheme);
	else writer.add(outTheme, theme ?? '', CONTENT_TYPES.theme);
	rels.add(RELATIONSHIP_TYPES.theme, relativeTarget(bookPart, outTheme));
	if (ctx.threaded.used) {
		const personsPart = personsSource ?? 'xl/persons/person.xml';
		writer.add(personsPart, ctx.persons.xml(), CONTENT_TYPES.persons);
		rels.add(RELATIONSHIP_TYPES.person, relativeTarget(bookPart, personsPart));
	}

	writeMetadataPart(
		writer,
		rels,
		bookPart,
		metadata,
		ctx.dynamicArrays?.used ?? false,
		metadataSource,
		reserved,
	);

	const sourceXml = sourceBook ? source?.text(sourceBook) : undefined;
	writer.add(
		bookPart,
		workbookXml(workbook, sheetRefs, Math.max(0, activeTab), sourceXml),
		keptVba ? CONTENT_TYPES.workbookMacro : CONTENT_TYPES.workbook,
	);
	writer.rels(bookPart, rels);

	const root = new RelationshipSet();
	root.add(RELATIONSHIP_TYPES.officeDocument, bookPart);
	root.add(RELATIONSHIP_TYPES.coreProperties, 'docProps/core.xml');
	root.add(RELATIONSHIP_TYPES.extendedProperties, 'docProps/app.xml');
	if (source) {
		for (const rel of source.rels('').values()) {
			if (ROOT_REGENERATED.has(rel.type) || regeneratesRootRel(workbook, rel.type)) continue;
			const target = source.target('', rel);
			// A regenerated package invalidates any digital signature (Excel drops it too), so the
			// origin, the signatures, their relationships and content types are never carried.
			if (rel.type === DIGITAL_SIGNATURE_ORIGIN_REL_TYPE || (target && isSignaturePart(target)))
				continue;
			if (target && target !== 'docProps/core.xml' && target !== 'docProps/app.xml') {
				root.add(rel.type, rel.target, rel.mode === 'External');
				writer.carry(target);
			}
		}
	}
	writeDocProps(writer, root, workbook, source);
	writer.rels('', root);
	// Carried parts (pivot caches, chartsheet charts) follow sheet renames and deletions.
	patchCarriedParts(writer, workbookRefEdits(workbook, source));
	return writer.build();
}
