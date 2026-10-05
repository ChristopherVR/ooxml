import { NS, first, type XmlElement } from '../../xml/index.js';
import type {
	CfvoThreshold,
	Color,
	ConditionalFormat,
	ConditionalOperator,
	ConditionalRule,
	DataValidation,
	DifferentialStyle,
	TimePeriod,
	ValidationType,
} from '../model.js';
import { TIME_PERIODS } from '../model.js';
import { stripFuturePrefixes } from './formula-text.js';
import { parseSqref } from './sheet-props.js';
import { parseColor } from './style-parts.js';
import { att, boolAttr, numAttr, xChildren, xFirst } from './xml-util.js';

const CFVO_TYPES = new Set(['min', 'max', 'num', 'percent', 'percentile', 'formula']);

function cfvo(node: XmlElement): CfvoThreshold {
	const type = att(node, 'type') ?? 'num';
	const threshold: CfvoThreshold = {
		type: (CFVO_TYPES.has(type) ? type : 'num') as CfvoThreshold['type'],
	};
	const value = att(node, 'val');
	if (value !== undefined) threshold.value = value;
	if (boolAttr(node, 'gte') === false) threshold.gte = false;
	return threshold;
}

const formulasOf = (rule: XmlElement) =>
	xChildren(rule, 'formula').map((f) => stripFuturePrefixes(f.textContent ?? ''));

const SIMPLE_TYPES = new Set([
	'duplicateValues',
	'uniqueValues',
	'containsBlanks',
	'notContainsBlanks',
	'containsErrors',
	'notContainsErrors',
]);
const TEXT_TYPES = new Set(['containsText', 'notContainsText', 'beginsWith', 'endsWith']);

function readRule(
	rule: XmlElement,
	dxfs: readonly DifferentialStyle[],
	palette: readonly string[] | undefined,
): ConditionalRule | undefined {
	const out = readRuleBody(rule, dxfs, palette);
	// `stopIfTrue` is an attribute of every `cfRule`, whatever its type.
	if (out && boolAttr(rule, 'stopIfTrue')) out.stopIfTrue = true;
	return out;
}

function readRuleBody(
	rule: XmlElement,
	dxfs: readonly DifferentialStyle[],
	palette: readonly string[] | undefined,
): ConditionalRule | undefined {
	const type = att(rule, 'type') ?? '';
	const priority = numAttr(rule, 'priority') ?? 1;
	const dxfId = numAttr(rule, 'dxfId');
	const style: DifferentialStyle = structuredClone(dxfId === undefined ? {} : (dxfs[dxfId] ?? {}));
	const colors = (parent: XmlElement | undefined): Color[] =>
		parent
			? xChildren(parent, 'color')
					.map((node) => parseColor(node, palette))
					.filter((c): c is Color => c !== undefined)
			: [];
	switch (type) {
		case 'cellIs':
			return {
				type,
				operator: (att(rule, 'operator') ?? 'equal') as ConditionalOperator,
				formulas: formulasOf(rule),
				style,
				priority,
			};
		case 'expression':
			return { type, formula: formulasOf(rule)[0] ?? '', style, priority };
		case 'colorScale': {
			const scale = xFirst(rule, 'colorScale');
			return {
				type,
				thresholds: scale ? xChildren(scale, 'cfvo').map(cfvo) : [],
				colors: colors(scale),
				priority,
			};
		}
		case 'dataBar': {
			const bar = xFirst(rule, 'dataBar');
			const [min = { type: 'min' }, max = { type: 'max' }] = bar
				? xChildren(bar, 'cfvo').map(cfvo)
				: [];
			const out: ConditionalRule = {
				type,
				min,
				max,
				color: colors(bar)[0] ?? { rgb: 'FF638EC6' },
				priority,
			};
			if (boolAttr(bar, 'showValue') === false) out.showValue = false;
			const ext = xFirst(rule, 'extLst');
			const id = ext ? first(xFirst(ext, 'ext'), 'id', NS.x14) : undefined;
			if (id?.textContent) out.extensionId = id.textContent.trim();
			return out;
		}
		case 'iconSet': {
			const set = xFirst(rule, 'iconSet');
			const out: ConditionalRule = {
				type,
				iconSet: att(set, 'iconSet') ?? '3TrafficLights1',
				thresholds: set ? xChildren(set, 'cfvo').map(cfvo) : [],
				priority,
			};
			if (boolAttr(set, 'reverse')) out.reverse = true;
			if (boolAttr(set, 'showValue') === false) out.showValue = false;
			return out;
		}
		case 'top10': {
			const out: ConditionalRule = { type, rank: numAttr(rule, 'rank') ?? 10, style, priority };
			if (boolAttr(rule, 'bottom')) out.bottom = true;
			if (boolAttr(rule, 'percent')) out.percent = true;
			return out;
		}
		case 'aboveAverage': {
			const out: ConditionalRule = { type, style, priority };
			if (boolAttr(rule, 'aboveAverage') === false) out.below = true;
			if (boolAttr(rule, 'equalAverage')) out.equalAverage = true;
			return out;
		}
		case 'timePeriod': {
			const period = att(rule, 'timePeriod') ?? '';
			if (!(TIME_PERIODS as readonly string[]).includes(period)) return undefined;
			return { type, timePeriod: period as TimePeriod, style, priority };
		}
		default:
			if (SIMPLE_TYPES.has(type)) return { type: type as 'duplicateValues', style, priority };
			if (TEXT_TYPES.has(type))
				return { type: type as 'containsText', text: att(rule, 'text') ?? '', style, priority };
			return undefined;
	}
}

/** Reads every `conditionalFormatting` block; unsupported rule types are reported and skipped. */
export function readConditionalFormats(
	blocks: readonly XmlElement[],
	dxfs: readonly DifferentialStyle[],
	palette: readonly string[] | undefined,
	warn: (message: string) => void,
): ConditionalFormat[] {
	const out: ConditionalFormat[] = [];
	for (const block of blocks) {
		const ranges = parseSqref(att(block, 'sqref'));
		const rules: ConditionalRule[] = [];
		for (const node of xChildren(block, 'cfRule')) {
			const rule = readRule(node, dxfs, palette);
			if (rule) rules.push(rule);
			else
				warn(
					`Conditional formatting rule type "${att(node, 'type') ?? ''}" is not supported and was dropped.`,
				);
		}
		if (ranges.length && rules.length) out.push({ ranges, rules });
	}
	return out;
}

const VALIDATION_TYPES = new Set([
	'none',
	'whole',
	'decimal',
	'list',
	'date',
	'time',
	'textLength',
	'custom',
]);
const ERROR_STYLES = new Set(['stop', 'warning', 'information']);

export function readDataValidations(node: XmlElement | undefined): DataValidation[] {
	if (!node) return [];
	const out: DataValidation[] = [];
	for (const dv of xChildren(node, 'dataValidation')) {
		const ranges = parseSqref(att(dv, 'sqref'));
		if (!ranges.length) continue;
		const type = att(dv, 'type') ?? 'none';
		const validation: DataValidation = {
			ranges,
			type: (VALIDATION_TYPES.has(type) ? type : 'none') as ValidationType,
		};
		const operator = att(dv, 'operator');
		if (operator) validation.operator = operator as ConditionalOperator;
		const f1 = xFirst(dv, 'formula1');
		if (f1) validation.formula1 = stripFuturePrefixes(f1.textContent ?? '');
		const f2 = xFirst(dv, 'formula2');
		if (f2) validation.formula2 = stripFuturePrefixes(f2.textContent ?? '');
		if (boolAttr(dv, 'allowBlank')) validation.allowBlank = true;
		const hideDropDown = boolAttr(dv, 'showDropDown');
		if (validation.type === 'list') validation.showDropDown = !hideDropDown;
		else if (hideDropDown) validation.showDropDown = false;
		if (boolAttr(dv, 'showErrorMessage')) validation.showErrorMessage = true;
		if (boolAttr(dv, 'showInputMessage')) validation.showInputMessage = true;
		const errorStyle = att(dv, 'errorStyle');
		if (errorStyle && ERROR_STYLES.has(errorStyle))
			validation.errorStyle = errorStyle as 'stop' | 'warning' | 'information';
		for (const key of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
			const value = att(dv, key);
			if (value !== undefined) validation[key] = value;
		}
		out.push(validation);
	}
	return out;
}
