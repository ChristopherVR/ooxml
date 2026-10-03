import { NS, elements, parseXml, type XmlElement } from '../../xml/index.js';
import { columnLabel, quoteSheetName } from '../address.js';
import type { DefinedName, Workbook } from '../model.js';
import { addFuturePrefixes } from '../read/formula-text.js';
import { selfContainedXml } from '../read/xml-util.js';
import { modernHashValues } from './password-hash.js';
import { XML_HEADER, attrs, escapeAttr, escapeText, inlineFragment } from './xml-out.js';

/** CT_Workbook children kept verbatim from the source, by position in the schema sequence. */
const KEPT_BEFORE_SHEETS = ['fileVersion', 'fileSharing'];
const KEPT_AFTER_SHEETS = ['functionGroups', 'externalReferences'];
const KEPT_TAIL = [
	'oleSize',
	'customWorkbookViews',
	'pivotCaches',
	'smartTagPr',
	'smartTagTypes',
	'webPublishing',
	'fileRecoveryPr',
	'webPublishObjects',
	'extLst',
];

export interface SheetRef {
	name: string;
	sheetId: number;
	state: string;
	relId: string;
}

const copyAttrs = (node: XmlElement | undefined, skip: readonly string[]) =>
	node
		? Array.from(node.attributes)
				.filter((a) => !skip.includes(a.name) && !a.name.startsWith('xmlns'))
				.map((a) => ` ${a.name}="${escapeAttr(a.value)}"`)
				.join('')
		: '';

function absoluteRange(
	range: NonNullable<NonNullable<Workbook['sheets'][number]['pageSetup']>['printArea']>,
): string {
	const cell = (row: number, col: number) => `$${columnLabel(col)}$${row + 1}`;
	return `${cell(range.start.row, range.start.col)}:${cell(range.end.row, range.end.col)}`;
}

/** Defined names plus print areas lifted onto page setup at load time. */
export function definedNamesOf(workbook: Workbook): DefinedName[] {
	const names = workbook.definedNames.filter(
		(name) =>
			name.name &&
			(name.localSheet === undefined ||
				(name.localSheet >= 0 && name.localSheet < workbook.sheets.length)),
	);
	workbook.sheets.forEach((sheet, index) => {
		const area = sheet.pageSetup?.printArea;
		if (!area || names.some((n) => n.name === '_xlnm.Print_Area' && n.localSheet === index)) return;
		names.push({
			name: '_xlnm.Print_Area',
			formula: `${quoteSheetName(sheet.name)}!${absoluteRange(area)}`,
			localSheet: index,
		});
	});
	return names;
}

/** Builds `workbook.xml`, carrying unmodelled elements of the source part. */
export function workbookXml(
	workbook: Workbook,
	sheets: readonly SheetRef[],
	activeTab: number,
	sourceXml: string | undefined,
): string {
	const kept = new Map<string, XmlElement>();
	if (sourceXml) {
		try {
			for (const node of elements(parseXml(sourceXml, { label: 'XLSX workbook' }).documentElement))
				if (node.localName && !kept.has(node.localName)) kept.set(node.localName, node);
		} catch {
			// An unreadable source workbook part contributes nothing.
		}
	}
	const raw = (names: readonly string[]) =>
		names
			.map((name) =>
				kept.has(name) ? inlineFragment(selfContainedXml(kept.get(name) as XmlElement)) : '',
			)
			.join('');
	let out = `${XML_HEADER}<workbook xmlns="${NS.x}" xmlns:r="${NS.r}">`;
	out +=
		raw(KEPT_BEFORE_SHEETS) ||
		'<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27328"/>';
	out += `<workbookPr${copyAttrs(kept.get('workbookPr'), ['date1904'])}${workbook.date1904 ? ' date1904="1"' : ''}/>`;
	if (workbook.structureLocked) {
		// The hashes come from the model only: a removed or changed password must not keep the
		// source's hash attributes.
		const skip = [
			'lockStructure',
			'workbookPassword',
			'workbookAlgorithmName',
			'workbookHashValue',
			'workbookSaltValue',
			'workbookSpinCount',
		];
		const modern = Object.entries(modernHashValues(workbook.workbookModernHash, 'workbook'))
			.filter((entry): entry is [string, string] => entry[1] !== undefined)
			.map(([name, value]) => ` ${name}="${escapeAttr(value)}"`)
			.join('');
		const password = workbook.workbookPasswordHash
			? ` workbookPassword="${escapeAttr(workbook.workbookPasswordHash)}"`
			: '';
		out += `<workbookProtection${copyAttrs(kept.get('workbookProtection'), skip)}${modern}${password} lockStructure="1"/>`;
	}
	const view = kept.get('bookViews') ? elements(kept.get('bookViews') as XmlElement)[0] : undefined;
	out += `<bookViews><workbookView${copyAttrs(view, ['activeTab', 'firstSheet', 'xr2:uid'])}${activeTab ? ` activeTab="${activeTab}"` : ''}/></bookViews>`;
	out += `<sheets>${sheets
		.map(
			(sheet) =>
				`<sheet${attrs({ name: sheet.name, sheetId: sheet.sheetId, state: sheet.state === 'visible' ? undefined : sheet.state })} r:id="${sheet.relId}"/>`,
		)
		.join('')}</sheets>`;
	out += raw(KEPT_AFTER_SHEETS);
	const names = definedNamesOf(workbook);
	if (names.length)
		out += `<definedNames>${names
			.map(
				(name) =>
					`<definedName${attrs({ name: name.name, comment: name.comment, localSheetId: name.localSheet, hidden: name.hidden || undefined })}>${escapeText(addFuturePrefixes(name.formula))}</definedName>`,
			)
			.join('')}</definedNames>`;
	const calc = kept.get('calcPr');
	const mode = workbook.calcMode === 'manual' ? ' calcMode="manual"' : '';
	out += `<calcPr${copyAttrs(calc, ['fullCalcOnLoad', 'calcId', 'calcMode'])} calcId="191029"${mode} fullCalcOnLoad="1"/>`;
	out += raw(KEPT_TAIL);
	return `${out}</workbook>`;
}

/** `docProps/core.xml` from the workbook properties. */
export function coreXml(workbook: Workbook): string {
	const p = workbook.properties;
	const text = (tag: string, value: string | undefined) =>
		value ? `<${tag}>${escapeText(value)}</${tag}>` : '';
	const date = (tag: string, value: string | undefined) =>
		value ? `<${tag} xsi:type="dcterms:W3CDTF">${escapeText(value)}</${tag}>` : '';
	return (
		`${XML_HEADER}<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
		text('dc:title', p.title) +
		text('dc:subject', p.subject) +
		text('dc:creator', p.creator) +
		text('cp:keywords', p.keywords) +
		text('dc:description', p.description) +
		text('cp:lastModifiedBy', p.lastModifiedBy) +
		date('dcterms:created', p.created) +
		date('dcterms:modified', p.modified) +
		'</cp:coreProperties>'
	);
}

/** `docProps/app.xml`: application, company and the sheet titles. */
export function appXml(workbook: Workbook): string {
	const names = workbook.sheets
		.map((sheet) => `<vt:lpstr>${escapeText(sheet.name)}</vt:lpstr>`)
		.join('');
	const company = workbook.properties.company
		? `<Company>${escapeText(workbook.properties.company)}</Company>`
		: '';
	return (
		`${XML_HEADER}<Properties xmlns="${NS.extendedProperties}" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
		`<Application>${escapeText(workbook.properties.application ?? 'Microsoft Excel')}</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>` +
		`<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${workbook.sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
		`<TitlesOfParts><vt:vector size="${workbook.sheets.length}" baseType="lpstr">${names}</vt:vector></TitlesOfParts>` +
		`${company}<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`
	);
}
