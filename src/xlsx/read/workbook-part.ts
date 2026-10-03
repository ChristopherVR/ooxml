import { NS, first, parseXml, relAttr } from '../../xml/index.js';
import type { DefinedName, ModernPasswordHash, SheetState, WorkbookProperties } from '../model.js';
import { stripFuturePrefixes } from './formula-text.js';
import { readModernHash } from './password-hash.js';
import { att, boolAttr, numAttr, xChildren, xFirst, xText } from './xml-util.js';

export interface SheetEntry {
	name: string;
	sheetId: number;
	state: SheetState;
	relId: string;
}

export interface WorkbookPart {
	sheets: SheetEntry[];
	definedNames: DefinedName[];
	date1904: boolean;
	activeTab: number;
	structureLocked: boolean;
	fullCalcOnLoad: boolean;
	/** `workbookProtection workbookPassword` (legacy hash, hex). */
	workbookPasswordHash?: string;
	/** `workbookProtection workbookAlgorithmName` and friends (Excel 2013+ agile hash). */
	workbookModernHash?: ModernPasswordHash;
	/** `calcPr calcMode`: `manual` (and `autoNoTable`, read as automatic). */
	calcMode?: 'auto' | 'manual';
}

export function parseWorkbookPart(xml: string): WorkbookPart {
	const root = parseXml(xml, { label: 'XLSX workbook' }).documentElement;
	const sheets: SheetEntry[] = [];
	xChildren(xFirst(root, 'sheets') ?? root, 'sheet').forEach((node, index) => {
		const state = att(node, 'state');
		sheets.push({
			name: att(node, 'name') ?? `Sheet${index + 1}`,
			sheetId: numAttr(node, 'sheetId') ?? index + 1,
			state: state === 'hidden' || state === 'veryHidden' ? state : 'visible',
			relId: relAttr(node, 'id') ?? '',
		});
	});
	const definedNames: DefinedName[] = [];
	for (const node of xChildren(xFirst(root, 'definedNames') ?? root, 'definedName')) {
		const name: DefinedName = {
			name: att(node, 'name') ?? '',
			formula: stripFuturePrefixes(xText(node)),
		};
		const local = numAttr(node, 'localSheetId');
		if (local !== undefined) name.localSheet = local;
		if (boolAttr(node, 'hidden')) name.hidden = true;
		const comment = att(node, 'comment');
		if (comment) name.comment = comment;
		if (name.name) definedNames.push(name);
	}
	const view = xFirst(xFirst(root, 'bookViews'), 'workbookView');
	const protection = xFirst(root, 'workbookProtection');
	const password = att(protection, 'workbookPassword');
	const modernHash = readModernHash(protection, 'workbook');
	const calcMode = att(xFirst(root, 'calcPr'), 'calcMode');
	return {
		...(password ? { workbookPasswordHash: password } : {}),
		...(modernHash ? { workbookModernHash: modernHash } : {}),
		...(calcMode === 'manual' ? { calcMode: 'manual' as const } : {}),
		sheets,
		definedNames,
		date1904: boolAttr(xFirst(root, 'workbookPr'), 'date1904', false),
		activeTab: numAttr(view, 'activeTab') ?? 0,
		structureLocked: boolAttr(protection, 'lockStructure', false),
		fullCalcOnLoad: boolAttr(xFirst(root, 'calcPr'), 'fullCalcOnLoad', false),
	};
}

/** Reads `docProps/core.xml` and `docProps/app.xml`. */
export function parseDocProps(
	core: string | undefined,
	app: string | undefined,
): WorkbookProperties {
	const props: WorkbookProperties = {};
	if (core) {
		const root = parseXml(core, { label: 'XLSX core properties' }).documentElement;
		const read = (ns: string, local: string, key: keyof WorkbookProperties) => {
			const node = first(root, local, ns);
			const text = node?.textContent ?? '';
			if (node && text !== '') props[key] = text;
		};
		read(NS.dc, 'title', 'title');
		read(NS.dc, 'subject', 'subject');
		read(NS.dc, 'creator', 'creator');
		read(NS.cp, 'keywords', 'keywords');
		read(NS.dc, 'description', 'description');
		read(NS.cp, 'lastModifiedBy', 'lastModifiedBy');
		read(NS.dcterms, 'created', 'created');
		read(NS.dcterms, 'modified', 'modified');
	}
	if (app) {
		const root = parseXml(app, { label: 'XLSX app properties' }).documentElement;
		const company = first(root, 'Company', NS.extendedProperties)?.textContent ?? '';
		if (company) props.company = company;
		const application = first(root, 'Application', NS.extendedProperties)?.textContent ?? '';
		if (application) props.application = application;
	}
	return props;
}
