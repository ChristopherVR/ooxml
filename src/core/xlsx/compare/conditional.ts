// Conditional formats have no identity of their own, and how a file groups rules into formats is
// an encoding detail: one format with two rules on A1:A9 paints like two formats with one rule
// each. So each rule is compared on its own, with the ranges of its format, matched by content
// (priority aside); neither renumbered priorities nor the order a file stores rules in is a change.
// The order rules are evaluated in is compared on its own.
import { formatAddress, formatRange } from '../address';
import type { CfvoThreshold, ConditionalFormat, ConditionalRule } from '../model';
import { stableKey } from '../stable-key';
import { withCanonicalColors } from './format-key';
import { longestIncreasing } from './sheets';
import type { SheetDifference } from './types';

/** One rule with the ranges of its format. */
interface Unit {
	format: ConditionalFormat;
	rule: ConditionalRule;
	key: string;
}

/** Rule settings that are off unless a file turns them on. */
export const OFF_BY_DEFAULT = [
	'stopIfTrue',
	'bottom',
	'percent',
	'below',
	'equalAverage',
	'reverse',
];

/** A threshold without `gte: true`, which is the default (a value equal to it reaches it). */
function threshold(t: CfvoThreshold): CfvoThreshold {
	if (t.gte !== true) return t;
	const { gte: _default, ...rest } = t;
	return rest;
}

const isEmptyObject = (value: unknown) =>
	!!value && typeof value === 'object' && Object.keys(value).length === 0;

/**
 * A rule without its priority, x14 extension ids (which differ between producers), spelled-out
 * defaults and empty parts of its differential style. That style keeps its `false` flags, which
 * switch a setting off.
 */
function ruleKey(rule: ConditionalRule): string {
	const out: Record<string, unknown> = {
		...rule,
		priority: undefined,
		extensionId: undefined,
		extensionXml: undefined,
	};
	for (const flag of OFF_BY_DEFAULT) if (out[flag] === false) delete out[flag];
	if (out.showValue === true) delete out.showValue;
	if ('thresholds' in rule) out.thresholds = rule.thresholds.map(threshold);
	if (rule.type === 'dataBar')
		Object.assign(out, { min: threshold(rule.min), max: threshold(rule.max) });
	if ('style' in rule)
		out.style = Object.fromEntries(Object.entries(rule.style).filter(([, v]) => !isEmptyObject(v)));
	return stableKey(withCanonicalColors(out));
}

/** The rule evaluates formulas, whose relative references the evaluator anchors at a range. */
const hasFormulas = (rule: ConditionalRule) =>
	rule.type === 'expression' ||
	rule.type === 'cellIs' ||
	('thresholds' in rule && rule.thresholds.some((t) => t.type === 'formula')) ||
	(rule.type === 'dataBar' && (rule.min.type === 'formula' || rule.max.type === 'formula'));

/**
 * Units keyed by sorted ranges and rule content. A rule with formulas also keeps the top-left cell
 * of its format's first range, where relative references are anchored (`cf-evaluator`), since
 * reordering the ranges then changes which cells the rule matches.
 */
const unitsOf = (formats: readonly ConditionalFormat[]): Unit[] =>
	formats.flatMap((format) => {
		const ranges = format.ranges.map(formatRange).sort();
		const start = format.ranges[0]?.start;
		const anchor = start ? formatAddress(start) : '';
		return format.rules.map((rule) => ({
			format,
			rule,
			key: stableKey([ranges, hasFormulas(rule) ? anchor : '', ruleKey(rule)]),
		}));
	});

/** Pairs equal rules, each occurrence once and equal occurrences in priority order. */
function pairUnits(before: readonly Unit[], after: readonly Unit[]) {
	const groups = new Map<string, { before: Unit[]; after: Unit[] }>();
	const groupOf = (unit: Unit) => {
		let group = groups.get(unit.key);
		if (!group) groups.set(unit.key, (group = { before: [], after: [] }));
		return group;
	};
	for (const unit of before) groupOf(unit).before.push(unit);
	for (const unit of after) groupOf(unit).after.push(unit);
	const pairedWith = new Map<Unit, Unit>();
	const byPriority = (x: Unit, y: Unit) => x.rule.priority - y.rule.priority;
	for (const group of groups.values()) {
		const next = group.after.sort(byPriority);
		group.before.sort(byPriority).forEach((unit, i) => {
			const match = next[i];
			if (match) pairedWith.set(unit, match);
		});
	}
	const paired = new Set(pairedWith.values());
	return {
		pairs: before.flatMap((unit) => {
			const match = pairedWith.get(unit);
			return match ? [{ before: unit, after: match }] : [];
		}),
		removed: before.filter((unit) => !pairedWith.has(unit)),
		added: after.filter((unit) => !paired.has(unit)),
	};
}

/**
 * The paired rules that moved above or below other rules. Priorities are compared by their order
 * among the rules the two sheets share, so adding a rule is not a change; a rule outside the
 * longest run that kept its order moved.
 */
function reordered<T extends { before: Unit; after: Unit }>(pairs: readonly T[]): T[] {
	const rankAfter = new Map(
		[...pairs].sort((x, y) => x.after.rule.priority - y.after.rule.priority).map((p, i) => [p, i]),
	);
	const byBefore = [...pairs].sort((x, y) => x.before.rule.priority - y.before.rule.priority);
	const kept = longestIncreasing(byBefore.map((pair) => rankAfter.get(pair) ?? 0));
	const moved = new Set(byBefore.filter((_, i) => !kept.has(i)));
	return pairs.filter((pair) => moved.has(pair));
}

/** A rule as its own one-rule format, as differences report it. */
const single = (unit: Unit): ConditionalFormat => ({
	ranges: unit.format.ranges,
	rules: [unit.rule],
});

/** Conditional format rules removed, added and reordered between two sheets. */
export function conditionalFormatDifferences(
	before: readonly ConditionalFormat[],
	after: readonly ConditionalFormat[],
): SheetDifference[] {
	const { pairs, removed, added } = pairUnits(unitsOf(before), unitsOf(after));
	return [
		...removed.map((unit) => ({ kind: 'conditionalFormatRemoved' as const, before: single(unit) })),
		...added.map((unit) => ({ kind: 'conditionalFormatAdded' as const, after: single(unit) })),
		...reordered(pairs).map((pair) => ({
			kind: 'conditionalFormatReordered' as const,
			before: single(pair.before),
			after: single(pair.after),
		})),
	];
}
