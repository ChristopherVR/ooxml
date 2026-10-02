// New / Edit Formatting Rule: pick a rule type, describe it, choose the format. `editRuleDialog`
// only returns the rule (the Rules Manager stages it); the 'cf-rule' dialog applies it.
import type { CellRange, ConditionalRule, DifferentialStyle } from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { customFormat, hasStyle, previewBox } from './cf-common.js';
import {
	type RulePanel,
	averagePanel,
	containsPanel,
	formulaPanel,
	topPanel,
	uniquePanel,
} from './cf-rule-panels.js';
import { valuesPanel } from './cf-rule-scale.js';
import { el, invalid, listBox, row, text } from './fields.js';
import { button, showDialog } from './frame.js';

export interface RuleEdit {
	rule: ConditionalRule;
	ranges: CellRange[];
}

type Kind = 'values' | 'contains' | 'top' | 'average' | 'unique' | 'formula';

const KINDS: ReadonlyArray<readonly [Kind, string]> = [
	['values', 'Format all cells based on their values'],
	['contains', 'Format only cells that contain'],
	['top', 'Format only top or bottom ranked values'],
	['average', 'Format only values that are above or below average'],
	['unique', 'Format only unique or duplicate values'],
	['formula', 'Use a formula to determine which cells to format'],
];

export function kindOf(rule: ConditionalRule): Kind {
	switch (rule.type) {
		case 'colorScale':
		case 'dataBar':
		case 'iconSet':
			return 'values';
		case 'top10':
			return 'top';
		case 'aboveAverage':
			return 'average';
		case 'duplicateValues':
		case 'uniqueValues':
			return 'unique';
		case 'expression':
			return 'formula';
		default:
			return 'contains';
	}
}

/** Shows the rule editor; resolves the edited rule (priority 1, the caller renumbers). */
export function editRuleDialog(
	ctx: EditorContext,
	initial?: RuleEdit,
): Promise<RuleEdit | undefined> {
	const ranges = initial?.ranges ?? target(ctx)?.ranges ?? [];
	const panels: Record<Kind, RulePanel> = {
		values: valuesPanel(ctx),
		contains: containsPanel(ctx),
		top: topPanel(ctx),
		average: averagePanel(ctx),
		unique: uniquePanel(ctx),
		formula: formulaPanel(ctx),
	};
	let style: DifferentialStyle =
		initial && hasStyle(initial.rule) ? structuredClone(initial.rule.style) : {};
	let kind: Kind = initial ? kindOf(initial.rule) : 'values';
	const preview = previewBox(ctx);
	const formatButton = button(ctx, 'Format...');
	const formatRow = row(ctx, preview.element, formatButton);
	const description = el(ctx, 'div', 'xve-cf-description');
	description.append(...Object.values(panels).map((p) => p.element), formatRow);
	const show = (next: Kind): void => {
		kind = next;
		for (const [k, p] of Object.entries(panels)) p.element.hidden = k !== next;
		formatRow.hidden = !panels[next].styled;
		preview.show(style);
	};
	const types = listBox(ctx, 'Select a Rule Type:', (value) => show(value as Kind));
	types.setItems(KINDS.map(([k, label]) => [k, ctx.t(label)] as const));
	if (initial) panels[kind].load(initial.rule);
	types.select(kind);
	show(kind);
	formatButton.addEventListener('click', async () => {
		const chosen = await customFormat(ctx, style);
		if (chosen) {
			style = chosen;
			preview.show(style);
		}
	});
	return showDialog<RuleEdit>(ctx, {
		name: 'cf-rule',
		heading: initial ? 'Edit Formatting Rule' : 'New Formatting Rule',
		wide: true,
		body: [
			text(ctx, 'Select a Rule Type:', 'xve-field-label'),
			types.element,
			text(ctx, 'Edit the Rule Description:', 'xve-field-label'),
			description,
		],
		opened: () => types.element.focus(),
		submit: () => {
			const read = panels[kind].read(structuredClone(style));
			if ('error' in read) return invalid(ctx, read.control, read.error, read.vars);
			const rule = keepExtras(read, initial?.rule);
			return { rule, ranges };
		},
	});
}

/** Keeps what the editor does not show (data bar extension id, stop-if-true, priority). */
function keepExtras(rule: ConditionalRule, original: ConditionalRule | undefined): ConditionalRule {
	if (!original) return rule;
	rule.priority = original.priority;
	if (rule.type === 'dataBar' && original.type === 'dataBar') {
		if (original.extensionId) rule.extensionId = original.extensionId;
		if (original.showValue !== undefined) rule.showValue = original.showValue;
	}
	const stops = (
		r: ConditionalRule,
	): r is Extract<ConditionalRule, { type: 'cellIs' | 'expression' | 'timePeriod' }> =>
		r.type === 'cellIs' || r.type === 'expression' || r.type === 'timePeriod';
	if (stops(rule) && stops(original) && original.stopIfTrue) rule.stopIfTrue = true;
	return rule;
}

/** The 'cf-rule' dialog: a new rule on the selection, or `{ format, rule }` edits that rule. */
export async function cfRuleDialog(
	ctx: EditorContext,
	props: unknown,
): Promise<RuleEdit | undefined> {
	const t = target(ctx);
	if (!t) return undefined;
	const at = props as { format?: number; rule?: number } | undefined;
	const format = at?.format !== undefined ? t.ws.conditionalFormats[at.format] : undefined;
	const existing = format && at?.rule !== undefined ? format.rules[at.rule] : undefined;
	if (format && existing && at?.format !== undefined && at.rule !== undefined) {
		const result = await editRuleDialog(ctx, { rule: existing, ranges: format.ranges });
		if (!result) return undefined;
		const next = structuredClone(format);
		next.rules[at.rule] = result.rule;
		t.session.replaceConditionalFormat(t.sheet, at.format, next);
		return result;
	}
	const result = await editRuleDialog(ctx);
	if (!result) return undefined;
	t.session.addConditionalFormat(t.sheet, { ranges: result.ranges, rules: [result.rule] });
	return result;
}
