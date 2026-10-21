// A plain-text report of a workbook comparison, for a terminal or a code review: one line per
// difference, in the result's stable order, with A1 addresses and old -> new values.
import { formatAddress, formatRange, type CellRange } from '../address';
import {
	isCellError,
	type CellValue,
	type ConditionalFormat,
	type DataValidation,
	type DefinedName,
} from '../model';
import { stableKey } from '../stable-key';
import { OFF_BY_DEFAULT } from './conditional';
import type {
	CellDifference,
	CellDifferenceRange,
	CellSide,
	FormatAttribute,
	SheetDifference,
	WorkbookComparison,
	WorkbookDifference,
} from './types';

export interface ComparisonTextOptions {
	/** One line per cell instead of one per range of neighbouring cells; default false. */
	cells?: boolean;
}

const KIND_TEXT: Record<CellDifference['kind'], string> = {
	enteredValue: 'entered value',
	formula: 'formula',
	calculatedValue: 'calculated value',
	format: 'format',
};

const ATTRIBUTE_TEXT: Record<FormatAttribute, string> = {
	numberFormat: 'number format',
	font: 'font',
	fill: 'fill',
	border: 'border',
	alignment: 'alignment',
	protection: 'protection',
	richText: 'rich text',
};

/** A quoted string, kept on one line (JSON leaves the line and paragraph separators literal). */
const quote = (text: string) => oneLine(JSON.stringify(text));

const ESCAPES: Record<string, string> = { '\n': '\\n', '\r': '\\r', '\t': '\\t' };

/**
 * Text kept on one line: line breaks, separators and other control characters escaped, so a
 * formula, a name or a value cannot start report lines of its own (Excel allows line breaks in
 * formulas).
 */
function oneLine(text: string): string {
	let out = '';
	for (const ch of text) {
		const code = ch.charCodeAt(0);
		// C0 and C1 controls (U+0085 is a line break to some readers) and the line and paragraph
		// separators U+2028 and U+2029.
		const control = code < 32 || (code >= 127 && code <= 159) || code === 0x2028 || code === 0x2029;
		out += control ? (ESCAPES[ch] ?? `\\u${code.toString(16).padStart(4, '0')}`) : ch;
	}
	return out;
}

/** A value as a report shows it: strings quoted, so `"1"` and `1` differ visibly. */
function valueText(value: CellValue): string {
	if (value === null) return '(empty)';
	if (typeof value === 'string') return quote(value);
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isCellError(value)) return value.error;
	return String(value);
}

/** A formula (with its array range) when the cell has one, else its value. */
function cellText(side: CellSide): string {
	if (side.formula === undefined) return valueText(side.value);
	const array = side.arrayRange ? ` (array ${formatRange(side.arrayRange)})` : '';
	const legacy = side.legacyFormula ? ' (implicit intersection)' : '';
	return `=${oneLine(side.formula)}${array}${legacy}`;
}

const position = (index: number) => `position ${index + 1}`;
const nameText = (name: string, scope: string | undefined) =>
	scope === undefined ? oneLine(name) : `${oneLine(name)} (${quote(scope)})`;
const definition = (name: DefinedName) =>
	`=${oneLine(name.formula)}${name.hidden ? ' (hidden)' : ''}${name.comment ? ` (comment ${quote(name.comment)})` : ''}`;

function nameChange(d: Extract<WorkbookDifference, { kind: 'nameChanged' }>): string {
	const from = definition(d.before);
	const to = definition(d.after);
	const renamed = d.before.name === d.after.name ? '' : ` renamed to ${oneLine(d.after.name)}`;
	const changed = from === to ? '' : `${renamed ? ' and' : ''} changed: ${from} -> ${to}`;
	return `Name ${nameText(d.before.name, d.scope)}${renamed}${changed}`;
}

function workbookLine(d: WorkbookDifference): string {
	switch (d.kind) {
		case 'sheetAdded':
			return `Sheet ${quote(d.sheet)} added at ${position(d.index)}`;
		case 'sheetRemoved':
			return `Sheet ${quote(d.sheet)} removed from ${position(d.index)}`;
		case 'sheetRenamed':
			return `Sheet ${quote(d.from)} renamed to ${quote(d.to)}`;
		case 'sheetMoved':
			return `Sheet ${quote(d.sheet)} moved from ${position(d.from)} to ${position(d.to)}`;
		case 'sheetState':
			return `Sheet ${quote(d.sheet)} ${d.from} -> ${d.to}`;
		case 'nameAdded':
			return `Name ${nameText(d.name, d.scope)} added: ${definition(d.after)}`;
		case 'nameRemoved':
			return `Name ${nameText(d.name, d.scope)} removed: ${definition(d.before)}`;
		case 'nameChanged':
			return nameChange(d);
	}
}

const rangesText = (list: readonly CellRange[]) => list.map(formatRange).join(' ');

/**
 * Flags that are off unless set, so `false` says nothing; any other `false` (a hidden drop-down,
 * a data bar without its value) is a setting and is shown.
 */
const UNSET = new Set([...OFF_BY_DEFAULT, 'allowBlank', 'showErrorMessage', 'showInputMessage']);

/** Settings shown first, in this order; the rest follow by name. */
const LEADING = ['type', 'operator', 'formula', 'formulas', 'formula1', 'formula2', 'text'];

/**
 * The settings of a rule or validation that are set, so two report lines that differ only in a
 * formula, a flag or a style say so: words bare, other strings quoted, structures as JSON.
 */
function settingsText(item: object, skip: readonly string[]): string {
	const rank = (key: string) => (LEADING.includes(key) ? LEADING.indexOf(key) : LEADING.length);
	const shown = (key: string, v: unknown) =>
		!skip.includes(key) && v !== undefined && v !== '' && !(v === false && UNSET.has(key));
	return Object.entries(item)
		.filter(([key, v]) => shown(key, v))
		.sort(([a], [b]) => rank(a) - rank(b) || (a < b ? -1 : a > b ? 1 : 0))
		.map(([key, v]) => {
			if (v === true) return key;
			if (v === false) return `${key} false`;
			if (typeof v !== 'string')
				return `${key} ${typeof v === 'object' ? stableKey(v) : String(v)}`;
			return `${key} ${/^[A-Za-z][\w.-]*$/.test(v) ? v : quote(v)}`;
		})
		.join(', ');
}

const validationText = (v: DataValidation) =>
	`${rangesText(v.ranges)} (${settingsText(v, ['ranges'])})`;

const ruleText = (format: ConditionalFormat) =>
	`${rangesText(format.ranges)} (${format.rules.map((rule) => settingsText(rule, ['priority', 'extensionId', 'extensionXml'])).join('; ')})`;

const priorities = (format: ConditionalFormat) =>
	format.rules.map((rule) => rule.priority).join(', ');

function sheetLine(d: SheetDifference): string {
	switch (d.kind) {
		case 'mergeAdded':
			return `Merge added ${formatRange(d.range)}`;
		case 'mergeRemoved':
			return `Merge removed ${formatRange(d.range)}`;
		case 'tableAdded':
			return `Table ${oneLine(d.table)} added at ${formatRange(d.after.range)}`;
		case 'tableRemoved':
			return `Table ${oneLine(d.table)} removed from ${formatRange(d.before.range)}`;
		case 'tableChanged': {
			const where =
				formatRange(d.before.range) === formatRange(d.after.range)
					? formatRange(d.after.range)
					: `${formatRange(d.before.range)} -> ${formatRange(d.after.range)}`;
			return `Table ${oneLine(d.table)} changed (${d.fields.join(', ')}): ${where}`;
		}
		case 'dataValidationAdded':
			return `Data validation added ${validationText(d.after)}`;
		case 'dataValidationRemoved':
			return `Data validation removed ${validationText(d.before)}`;
		case 'conditionalFormatAdded':
			return `Conditional format rule added ${ruleText(d.after)}`;
		case 'conditionalFormatRemoved':
			return `Conditional format rule removed ${ruleText(d.before)}`;
		case 'conditionalFormatReordered':
			return `Conditional format rule reordered ${ruleText(d.after)}: priority ${priorities(d.before)} -> ${priorities(d.after)}`;
	}
}

function attributesText(attributes: readonly FormatAttribute[] | undefined): string {
	return (attributes ?? []).map((a) => ATTRIBUTE_TEXT[a]).join(', ');
}

function cellLine(d: CellDifference): string {
	const at = `${formatAddress(d.address)} ${KIND_TEXT[d.kind]}`;
	if (d.kind === 'formula') return `${at}: ${cellText(d.before)} -> ${cellText(d.after)}`;
	if (d.kind !== 'format') {
		const formula = d.after.formula === undefined ? '' : ` (=${oneLine(d.after.formula)})`;
		return `${at}: ${valueText(d.before.value)} -> ${valueText(d.after.value)}${formula}`;
	}
	const line = `${at} (${attributesText(d.attributes)})`;
	const from = d.before.style?.numFmt;
	const to = d.after.style?.numFmt;
	return d.attributes?.includes('numberFormat') && from !== undefined && to !== undefined
		? `${line}: ${quote(from)} -> ${quote(to)}`
		: line;
}

function rangeLine(range: CellDifferenceRange, first: CellDifference | undefined): string {
	if (range.count === 1 && first) return cellLine(first);
	const attributes = range.attributes ? ` (${attributesText(range.attributes)})` : '';
	return `${formatRange(range.range)} ${KIND_TEXT[range.kind]}${attributes}: ${range.count} cells`;
}

/** The comparison as text: workbook changes, then each sheet with differences, then the limits. */
export function comparisonText(
	result: WorkbookComparison,
	options: ComparisonTextOptions = {},
): string {
	const lines: string[] = [];
	if (result.workbook.length) {
		lines.push('Workbook');
		for (const d of result.workbook) lines.push(`  ${workbookLine(d)}`);
	}
	for (const sheet of result.sheets) {
		if (!sheet.structure.length && !sheet.cells.length) continue;
		const renamed = sheet.beforeName === sheet.name ? '' : ` (was ${quote(sheet.beforeName)})`;
		lines.push(`Sheet ${quote(sheet.name)}${renamed}`);
		for (const d of sheet.structure) lines.push(`  ${sheetLine(d)}`);
		if (options.cells) {
			for (const d of sheet.cells) lines.push(`  ${cellLine(d)}`);
		} else {
			const firstAt = new Map<string, CellDifference>();
			for (const d of sheet.cells) firstAt.set(`${d.kind}\n${formatAddress(d.address)}`, d);
			for (const range of sheet.ranges) {
				const first = firstAt.get(`${range.kind}\n${formatAddress(range.range.start)}`);
				lines.push(`  ${rangeLine(range, first)}`);
			}
		}
	}
	if (!lines.length) lines.push('No differences.');
	if (result.truncated) lines.push(`Stopped after ${result.count} differences; there may be more.`);
	lines.push(`Not compared: ${result.notCompared.join(', ')}.`);
	return `${lines.join('\n')}\n`;
}
