import { NS, elements, parseXml, type XmlElement } from '../../xml/index';
import { columnLabel, quoteSheetName } from '../address';
import type { DefinedName, Workbook } from '../model';
import { addFuturePrefixes } from '../read/formula-text';
import { selfContainedXml } from '../read/xml-util';
import { modernHashValues } from './password-hash';
import { XML_HEADER, attrs, escapeAttr, escapeText, inlineFragment } from './xml-out';

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
	pivotCachesXml?: string,
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
	const views = kept.get('bookViews') ? elements(kept.get('bookViews') as XmlElement) : [];
	// The first window follows the model's active sheet; further windows are kept verbatim.
	const first = views[0];
	const otherWindows = views
		.slice(1)
		.map((view) => inlineFragment(selfContainedXml(view)))
		.join('');
	out += `<bookViews><workbookView${copyAttrs(first, ['activeTab', 'firstSheet', 'xr2:uid'])}${activeTab ? ` activeTab="${activeTab}"` : ''}/>${otherWindows}</bookViews>`;
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
					`<definedName${attrs({ name: name.name, comment: name.comment, localSheetId: name.localSheet, hidden: name.hidden || undefined, ...name.extra })}>${escapeText(addFuturePrefixes(name.formula))}</definedName>`,
			)
			.join('')}</definedNames>`;
	const calc = kept.get('calcPr');
	const mode = workbook.calcMode === 'manual' ? ' calcMode="manual"' : '';
	const iterate = workbook.iterate
		? ` iterate="1" iterateCount="${workbook.iterate.count}" iterateDelta="${workbook.iterate.delta}"`
		: '';
	const owned = ['fullCalcOnLoad', 'calcId', 'calcMode', 'iterate', 'iterateCount', 'iterateDelta'];
	out += `<calcPr${copyAttrs(calc, owned)} calcId="191029"${mode}${iterate} fullCalcOnLoad="1"/>`;
	// The model owns oleSize; a source's own element is kept only when the model sets none.
	const tailOf = (names: readonly string[]) =>
		names
			.map((name) =>
				name === 'oleSize' && workbook.oleSize
					? `<oleSize ref="${escapeAttr(workbook.oleSize)}"/>`
					: raw([name]),
			)
			.join('');
	out +=
		pivotCachesXml === undefined
			? tailOf(KEPT_TAIL)
			: tailOf(KEPT_TAIL.slice(0, 2)) + pivotCachesXml + tailOf(KEPT_TAIL.slice(3));
	return `${out}</workbook>`;
}

export { appXml, coreXml } from './doc-props';
