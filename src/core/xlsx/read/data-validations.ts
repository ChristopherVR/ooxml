import { NS, children, first, type XmlElement } from '../../xml/index';
import type { ConditionalOperator, DataValidation, ValidationType } from '../model';
import { stripFuturePrefixes } from './formula-text';
import { parseSqref } from './sheet-props';
import { att, boolAttr, selfContainedXml } from './xml-util';

export const VALIDATION_XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const TYPES = new Set(['none', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom']);
const ERROR_STYLES = new Set(['stop', 'warning', 'information']);

/** The same validation model for the original and Excel 2010 worksheet codecs. */
export function readDataValidation(dv: XmlElement): DataValidation | undefined {
	const extended = dv.namespaceURI === NS.x14;
	const ns = extended ? NS.x14 : NS.x;
	const ranges = parseSqref(
		extended ? (first(dv, 'sqref', VALIDATION_XM)?.textContent ?? '') : att(dv, 'sqref'),
	);
	if (!ranges.length) return undefined;
	const type = att(dv, 'type') ?? 'none';
	const validation: DataValidation = {
		ranges,
		type: (TYPES.has(type) ? type : 'none') as ValidationType,
	};
	const operator = att(dv, 'operator');
	if (operator) validation.operator = operator as ConditionalOperator;
	for (const key of ['formula1', 'formula2'] as const) {
		const container = first(dv, key, ns);
		const formula = extended ? first(container, 'f', VALIDATION_XM) : container;
		if (formula) validation[key] = stripFuturePrefixes(formula.textContent ?? '');
	}
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
	if (extended) validation.extensionXml = selfContainedXml(dv);
	return validation;
}

export function readDataValidations(node: XmlElement | undefined): DataValidation[] {
	return node
		? children(node, 'dataValidation', NS.x).flatMap((dv) => readDataValidation(dv) ?? [])
		: [];
}
