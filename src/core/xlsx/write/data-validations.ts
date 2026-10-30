import { NS, buildXml, first, parseXml } from '../../xml/index';
import { formatRange } from '../address';
import { tokenize } from '../formula/tokenizer';
import type { DataValidation } from '../model';
import { VALIDATION_XM } from '../read/data-validations';
import { addFuturePrefixes } from '../read/formula-text';
import { attrs, escapeText } from './xml-out';

/** Direct sheet references require the Excel 2010 form; quoted list items do not. */
export function usesValidationExtension(dv: DataValidation): boolean {
	if (dv.extensionXml) return true;
	return [dv.formula1, dv.formula2].some((formula) => {
		if (!formula) return false;
		try {
			return tokenize(formula).some((token) => token.prefix !== undefined);
		} catch {
			return false;
		}
	});
}

function validationAttributes(dv: DataValidation) {
	return {
		type: dv.type === 'none' ? undefined : dv.type,
		errorStyle: dv.errorStyle,
		operator: dv.operator,
		allowBlank: dv.allowBlank || undefined,
		showDropDown: dv.showDropDown === false ? true : undefined,
		showInputMessage: dv.showInputMessage || undefined,
		showErrorMessage: dv.showErrorMessage || undefined,
		errorTitle: dv.errorTitle,
		error: dv.error,
		promptTitle: dv.promptTitle,
		prompt: dv.prompt,
	};
}

export function dataValidationXml(dv: DataValidation, extended: boolean): string {
	const sqref = dv.ranges.map(formatRange).join(' ');
	if (!extended) {
		const formulas = (['formula1', 'formula2'] as const)
			.map((key) =>
				dv[key] === undefined ? '' : `<${key}>${escapeText(addFuturePrefixes(dv[key]))}</${key}>`,
			)
			.join('');
		const values = attrs({ ...validationAttributes(dv), sqref });
		return formulas
			? `<dataValidation${values}>${formulas}</dataValidation>`
			: `<dataValidation${values}/>`;
	}
	const doc = parseXml(
		dv.extensionXml || `<x14:dataValidation xmlns:x14="${NS.x14}" xmlns:xm="${VALIDATION_XM}"/>`,
	);
	const node = doc.documentElement;
	for (const [key, value] of Object.entries(validationAttributes(dv))) {
		if (value === undefined) node.removeAttribute(key);
		else node.setAttribute(key, typeof value === 'boolean' ? '1' : value);
	}
	for (const key of ['formula1', 'formula2'] as const) {
		let container = first(node, key, NS.x14);
		if (dv[key] === undefined) {
			if (container) node.removeChild(container);
			continue;
		}
		if (!container) {
			container = doc.createElementNS(NS.x14, `x14:${key}`);
			const next = key === 'formula1' ? first(node, 'formula2', NS.x14) : undefined;
			node.insertBefore(container, next ?? first(node, 'sqref', VALIDATION_XM) ?? null);
		}
		let formula = first(container, 'f', VALIDATION_XM);
		if (!formula) {
			formula = doc.createElementNS(VALIDATION_XM, 'xm:f');
			container.appendChild(formula);
		}
		formula.textContent = addFuturePrefixes(dv[key]);
	}
	let ranges = first(node, 'sqref', VALIDATION_XM);
	if (!ranges) {
		ranges = doc.createElementNS(VALIDATION_XM, 'xm:sqref');
		node.appendChild(ranges);
	}
	ranges.textContent = sqref;
	return buildXml(doc);
}

export function dataValidationsXml(validations: readonly DataValidation[]): string {
	const items = validations
		.filter((dv) => dv.ranges.length && !usesValidationExtension(dv))
		.map((dv) => dataValidationXml(dv, false));
	return items.length
		? `<dataValidations count="${items.length}">${items.join('')}</dataValidations>`
		: '';
}
