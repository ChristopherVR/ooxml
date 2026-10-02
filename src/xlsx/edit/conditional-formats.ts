// Editing existing conditional formats: replace, remove and reorder rules by priority.
import { normalizeRange } from '../address.js';
import type { ConditionalFormat, ConditionalRule, Worksheet } from '../model.js';
import { type EditContext, sheetAt } from './context.js';

function formatAt(sheet: Worksheet, index: number): ConditionalFormat {
	const format = sheet.conditionalFormats[index];
	if (!format) throw new RangeError(`No conditional format at index ${index}`);
	return format;
}

/** Every rule of the sheet, highest priority (lowest number) first; ties keep sheet order. */
function rulesByPriority(sheet: Worksheet): ConditionalRule[] {
	return sheet.conditionalFormats
		.flatMap((f) => f.rules)
		.map((rule, order) => ({ rule, order }))
		.sort((a, b) => a.rule.priority - b.rule.priority || a.order - b.order)
		.map((e) => e.rule);
}

/** Renumbers the rules 1..n in the given order. */
function renumber(rules: readonly ConditionalRule[]): void {
	rules.forEach((rule, i) => {
		rule.priority = i + 1;
	});
}

/**
 * Replaces the conditional format at `index` (ranges and rules). The new rules keep the
 * priorities they carry; rules are then renumbered 1..n so priorities stay unique.
 */
export function replaceConditionalFormat(
	ctx: EditContext,
	s: number,
	index: number,
	format: ConditionalFormat,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	formatAt(sheet, index);
	if (!format.ranges.length || !format.rules.length)
		throw new Error('A conditional format needs at least one range and one rule.');
	ctx.run(
		'Edit rule',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const next = structuredClone(format);
			next.ranges = next.ranges.map(normalizeRange);
			sheet.conditionalFormats[index] = next;
			renumber(rulesByPriority(sheet));
		},
		{ sheet: s, ranges: format.ranges },
	);
}

/** Removes the conditional format at `index` with all its rules. */
export function removeConditionalFormat(ctx: EditContext, s: number, index: number): void {
	const sheet = sheetAt(ctx.workbook, s);
	const format = formatAt(sheet, index);
	ctx.run(
		'Delete rule',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.conditionalFormats.splice(index, 1);
			renumber(rulesByPriority(sheet));
		},
		{ sheet: s, ranges: format.ranges },
	);
}

/**
 * Gives one rule a new priority (1 is evaluated first), shifting the others like the Rules
 * Manager's arrows; every rule of the sheet ends up numbered 1..n.
 */
export function setConditionalRulePriority(
	ctx: EditContext,
	s: number,
	formatIndex: number,
	ruleIndex: number,
	priority: number,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const format = formatAt(sheet, formatIndex);
	const rule = format.rules[ruleIndex];
	if (!rule) throw new RangeError(`No rule at index ${ruleIndex}`);
	const order = rulesByPriority(sheet);
	const from = order.indexOf(rule);
	const to = Math.max(0, Math.min(order.length - 1, Math.round(priority) - 1));
	if (from === to && order.every((r, i) => r.priority === i + 1)) return;
	ctx.run(
		'Rule priority',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const live = rulesByPriority(sheet);
			const target = sheet.conditionalFormats[formatIndex]?.rules[ruleIndex];
			if (!target) return;
			live.splice(live.indexOf(target), 1);
			live.splice(to, 0, target);
			renumber(live);
		},
		{ sheet: s, ranges: format.ranges },
	);
}

/** Moves a rule one step up (earlier) or down (later) in the evaluation order. */
export function moveConditionalRule(
	ctx: EditContext,
	s: number,
	formatIndex: number,
	ruleIndex: number,
	direction: 'up' | 'down',
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const rule = formatAt(sheet, formatIndex).rules[ruleIndex];
	if (!rule) throw new RangeError(`No rule at index ${ruleIndex}`);
	const position = rulesByPriority(sheet).indexOf(rule) + 1;
	setConditionalRulePriority(
		ctx,
		s,
		formatIndex,
		ruleIndex,
		direction === 'up' ? position - 1 : position + 1,
	);
}
