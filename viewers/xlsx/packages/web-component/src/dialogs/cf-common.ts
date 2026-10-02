// Shared pieces of the conditional formatting dialogs: operand quoting, "Applies to" parsing,
// rule descriptions, the format preview and the Custom Format bridge to the Format Cells dialog.
import {
	type CellRange,
	type ConditionalOperator,
	type ConditionalRule,
	type DifferentialStyle,
	type TimePeriod,
	formatRange,
	parseRange,
} from '@christophervr/xlsx-core';
import { dxfCss } from '../commands/cf-presets.js';
import type { EditorContext } from '../context.js';
import { el } from './fields.js';

/** A rule operand as typed: `=...` is a formula, a number stays, other text becomes a literal. */
export function operand(text: string): string {
	const value = text.trim();
	if (value.startsWith('=')) return value.slice(1);
	if (value !== '' && Number.isFinite(Number(value))) return value;
	if (/^".*"$/.test(value)) return value;
	return `"${value.replace(/"/g, '""')}"`;
}

/** The text a stored operand shows in a field (formulas get their `=` back). */
export function operandText(formula: string | undefined): string {
	if (formula === undefined) return '';
	if (Number.isFinite(Number(formula)) && formula.trim() !== '') return formula;
	const quoted = /^"(.*)"$/.exec(formula);
	if (quoted) return (quoted[1] ?? '').replace(/""/g, '"');
	return `=${formula}`;
}

/** Parses "A1:B2 C3" or "=$A$1:$B$2,C3" into ranges; undefined when any part is invalid. */
export function parseRefs(text: string): CellRange[] | undefined {
	const parts = text
		.replace(/^=/, '')
		.split(/[\s,;]+/)
		.map((p) => p.replace(/^.*!/, '').replace(/\$/g, ''))
		.filter(Boolean);
	if (!parts.length) return undefined;
	const out: CellRange[] = [];
	for (const part of parts) {
		const r = parseRange(part);
		if (!r) return undefined;
		out.push(r);
	}
	return out;
}

export const formatRefs = (ranges: CellRange[]): string =>
	ranges.map((r) => formatRange(r).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')).join(' ');

export const OPERATORS: ReadonlyArray<readonly [ConditionalOperator, string]> = [
	['between', 'between'],
	['notBetween', 'not between'],
	['equal', 'equal to'],
	['notEqual', 'not equal to'],
	['greaterThan', 'greater than'],
	['lessThan', 'less than'],
	['greaterThanOrEqual', 'greater than or equal to'],
	['lessThanOrEqual', 'less than or equal to'],
];

/** "A Date Occurring" periods in the order Excel's drop-down lists them. */
export const TIME_PERIOD_LABELS: ReadonlyArray<readonly [TimePeriod, string]> = [
	['yesterday', 'Yesterday'],
	['today', 'Today'],
	['tomorrow', 'Tomorrow'],
	['last7Days', 'In the last 7 days'],
	['lastWeek', 'Last week'],
	['thisWeek', 'This week'],
	['nextWeek', 'Next week'],
	['lastMonth', 'Last month'],
	['thisMonth', 'This month'],
	['nextMonth', 'Next month'],
];

/** A one-line description of a rule, as the Rules Manager lists it. */
export function describeRule(ctx: EditorContext, rule: ConditionalRule): string {
	const t = ctx.t;
	switch (rule.type) {
		case 'cellIs': {
			const op = OPERATORS.find(([o]) => o === rule.operator)?.[1] ?? rule.operator;
			const values = rule.formulas.map(operandText).join(` ${t('and')} `);
			return `${t('Cell Value')} ${t(op)} ${values}`;
		}
		case 'expression':
			return `${t('Formula:')} =${rule.formula}`;
		case 'colorScale':
			return t(rule.colors.length === 3 ? 'Graded Color Scale (3)' : 'Graded Color Scale (2)');
		case 'dataBar':
			return t('Data Bar');
		case 'iconSet':
			return `${t('Icon Set')} (${rule.iconSet})`;
		case 'top10':
			return t(
				rule.bottom
					? rule.percent
						? 'Bottom {n}%'
						: 'Bottom {n}'
					: rule.percent
						? 'Top {n}%'
						: 'Top {n}',
				{ n: rule.rank },
			);
		case 'aboveAverage':
			return t(
				rule.below
					? rule.equalAverage
						? 'Equal to or Below Average'
						: 'Below Average'
					: rule.equalAverage
						? 'Equal to or Above Average'
						: 'Above Average',
			);
		case 'duplicateValues':
			return t('Duplicate Values');
		case 'uniqueValues':
			return t('Unique Values');
		case 'containsBlanks':
			return t('Cell contains a blank value');
		case 'notContainsBlanks':
			return t('Cell does not contain a blank value');
		case 'containsErrors':
			return t('Cell contains an error');
		case 'notContainsErrors':
			return t('Cell does not contain an error');
		case 'containsText':
			return t('Cell Value contains "{text}"', { text: rule.text });
		case 'notContainsText':
			return t('Cell Value does not contain "{text}"', { text: rule.text });
		case 'beginsWith':
			return t('Cell Value begins with "{text}"', { text: rule.text });
		case 'endsWith':
			return t('Cell Value ends with "{text}"', { text: rule.text });
		case 'timePeriod': {
			const label = TIME_PERIOD_LABELS.find(([p]) => p === rule.timePeriod)?.[1];
			return `${t('Dates Occurring')}: ${t(label ?? rule.timePeriod)}`;
		}
	}
}

/** The "AaBbCcYyZz" sample of a differential style (or a gradient for scales and bars). */
export function previewBox(ctx: EditorContext): {
	element: HTMLDivElement;
	show(rule: ConditionalRule | DifferentialStyle | undefined): void;
} {
	const box = el(ctx, 'div', 'xve-sample xve-cf-preview');
	box.setAttribute('aria-label', ctx.t('Preview'));
	box.textContent = ctx.t('AaBbCcYyZz');
	return {
		element: box,
		show(input) {
			box.removeAttribute('style');
			if (!input) return;
			if ('type' in input) {
				const rule = input as ConditionalRule;
				if (rule.type === 'colorScale') {
					const stops = rule.colors.map((c) => `#${(c.rgb ?? 'FFFFFF').slice(-6)}`).join(',');
					box.setAttribute('style', `background:linear-gradient(90deg,${stops})`);
					return;
				}
				if (rule.type === 'dataBar') {
					const c = `#${(rule.color.rgb ?? '638EC6').slice(-6)}`;
					box.setAttribute('style', `background:linear-gradient(90deg,${c},#ffffff)`);
					return;
				}
				if (rule.type === 'iconSet') return;
				box.setAttribute('style', dxfCss(rule.style));
				return;
			}
			box.setAttribute('style', dxfCss(input as DifferentialStyle));
		},
	};
}

/** Opens Format Cells in differential mode; resolves the new style or undefined. */
export async function customFormat(
	ctx: EditorContext,
	dxf: DifferentialStyle,
): Promise<DifferentialStyle | undefined> {
	const result = await ctx.dialogs.open<DifferentialStyle>('format-cells', {
		dxf: structuredClone(dxf),
	});
	return result && typeof result === 'object' ? result : undefined;
}

/** Whether a rule carries a differential style. */
export const hasStyle = (
	rule: ConditionalRule,
): rule is Extract<ConditionalRule, { style: DifferentialStyle }> => 'style' in rule;

/** Whether a rule type supports Stop If True in the model. */
export const hasStopIfTrue = (
	rule: ConditionalRule,
): rule is Extract<ConditionalRule, { stopIfTrue?: boolean }> =>
	rule.type === 'cellIs' || rule.type === 'expression';
