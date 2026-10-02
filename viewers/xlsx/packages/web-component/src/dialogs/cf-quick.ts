// Home > Conditional Formatting > Highlight Cells Rules and Top/Bottom Rules: Excel's small
// one-line rule dialogs ("Format cells that are GREATER THAN: [value] with [format]").
import type { ConditionalRule, DifferentialStyle, TimePeriod } from '@christophervr/xlsx-core';
import { HIGHLIGHT_STYLES } from '../commands/cf-presets.js';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { TIME_PERIOD_LABELS, customFormat, operand } from './cf-common.js';
import { field, invalid, numberInput, row, select, text, textInput } from './fields.js';
import { showDialog } from './frame.js';

interface QuickSpec {
	heading: string;
	prompt: string;
	inputs: 0 | 1 | 2;
	rank?: boolean;
}

const SPECS: Readonly<Record<string, QuickSpec>> = {
	greaterThan: {
		heading: 'Greater Than',
		prompt: 'Format cells that are GREATER THAN:',
		inputs: 1,
	},
	lessThan: { heading: 'Less Than', prompt: 'Format cells that are LESS THAN:', inputs: 1 },
	between: { heading: 'Between', prompt: 'Format cells that are BETWEEN:', inputs: 2 },
	equal: { heading: 'Equal To', prompt: 'Format cells that are EQUAL TO:', inputs: 1 },
	containsText: {
		heading: 'Text That Contains',
		prompt: 'Format cells that contain the text:',
		inputs: 1,
	},
	timePeriod: {
		heading: 'A Date Occurring',
		prompt: 'Format cells that contain a date occurring:',
		inputs: 0,
	},
	duplicateValues: { heading: 'Duplicate Values', prompt: 'Format cells that contain:', inputs: 0 },
	top10: {
		heading: 'Top 10 Items',
		prompt: 'Format cells that rank in the TOP:',
		inputs: 0,
		rank: true,
	},
	top10Percent: {
		heading: 'Top 10%',
		prompt: 'Format cells that rank in the TOP:',
		inputs: 0,
		rank: true,
	},
	bottom10: {
		heading: 'Bottom 10 Items',
		prompt: 'Format cells that rank in the BOTTOM:',
		inputs: 0,
		rank: true,
	},
	bottom10Percent: {
		heading: 'Bottom 10%',
		prompt: 'Format cells that rank in the BOTTOM:',
		inputs: 0,
		rank: true,
	},
	aboveAverage: {
		heading: 'Above Average',
		prompt: 'Format cells that are ABOVE AVERAGE:',
		inputs: 0,
	},
	belowAverage: {
		heading: 'Below Average',
		prompt: 'Format cells that are BELOW AVERAGE:',
		inputs: 0,
	},
};

export const QUICK_KINDS = Object.keys(SPECS);

/** The rule a quick dialog builds from its fields (exported for tests). */
export function quickRule(
	kind: string,
	values: string[],
	style: DifferentialStyle,
	extra: { rank?: number; unique?: boolean; period?: TimePeriod },
): ConditionalRule {
	const priority = 1;
	switch (kind) {
		case 'greaterThan':
		case 'lessThan':
		case 'equal':
			return {
				type: 'cellIs',
				operator: kind,
				formulas: [operand(values[0] ?? '')],
				style,
				priority,
			};
		case 'between':
			return {
				type: 'cellIs',
				operator: 'between',
				formulas: [operand(values[0] ?? ''), operand(values[1] ?? '')],
				style,
				priority,
			};
		case 'containsText':
			return { type: 'containsText', text: values[0] ?? '', style, priority };
		case 'timePeriod':
			return { type: 'timePeriod', timePeriod: extra.period ?? 'today', style, priority };
		case 'duplicateValues':
			return { type: extra.unique ? 'uniqueValues' : 'duplicateValues', style, priority };
		case 'aboveAverage':
		case 'belowAverage':
			return {
				type: 'aboveAverage',
				...(kind === 'belowAverage' ? { below: true } : {}),
				style,
				priority,
			};
		default: {
			const rule: ConditionalRule = { type: 'top10', rank: extra.rank ?? 10, style, priority };
			if (kind.startsWith('bottom')) rule.bottom = true;
			if (kind.endsWith('Percent')) rule.percent = true;
			return rule;
		}
	}
}

export function cfQuickDialog(
	ctx: EditorContext,
	props: unknown,
): Promise<ConditionalRule | undefined> {
	const t = target(ctx);
	const kind = (props as { kind?: string } | undefined)?.kind ?? 'greaterThan';
	const spec = SPECS[kind];
	if (!t || !spec) return Promise.resolve(undefined);
	const inputs = Array.from({ length: spec.inputs }, () => textInput(ctx));
	const rank = numberInput(ctx, 10, 1, 1000);
	const dupSelect = select(
		ctx,
		[
			['duplicate', 'Duplicate'],
			['unique', 'Unique'],
		],
		'duplicate',
	);
	const period = select(ctx, TIME_PERIOD_LABELS, 'yesterday');
	const options: [string, string][] = HIGHLIGHT_STYLES.map(([id, label]) => [id, label]);
	options.push(['custom', 'Custom Format...']);
	const withSelect = select(ctx, options, 'lightRed');
	let custom: DifferentialStyle | undefined;
	let previous = 'lightRed';
	withSelect.addEventListener('change', async () => {
		if (withSelect.value !== 'custom') {
			previous = withSelect.value;
			return;
		}
		const base = custom ?? HIGHLIGHT_STYLES.find(([id]) => id === previous)?.[2] ?? {};
		const chosen = await customFormat(ctx, base);
		if (chosen) custom = chosen;
		else withSelect.value = previous;
	});
	const controls: HTMLElement[] = [];
	if (spec.rank) controls.push(field(ctx, spec.prompt.endsWith('TOP:') ? 'Top' : 'Bottom', rank));
	if (kind === 'duplicateValues') controls.push(field(ctx, 'Values', dupSelect));
	if (kind === 'timePeriod') controls.push(field(ctx, 'Date', period));
	inputs.forEach((input, i) =>
		controls.push(field(ctx, i === 0 ? (spec.inputs === 2 ? 'From' : 'Value') : 'and', input)),
	);
	const body = [
		text(ctx, spec.prompt, 'xve-field-label'),
		row(ctx, ...controls),
		field(ctx, kind.endsWith('Percent') ? '% with' : 'with', withSelect),
	];
	return showDialog<ConditionalRule>(ctx, {
		name: 'cf-quick',
		heading: spec.heading,
		body,
		opened: () =>
			(inputs[0] ?? (spec.rank ? rank : kind === 'timePeriod' ? period : withSelect)).focus(),
		submit: () => {
			const empty = inputs.find((i) => !i.value.trim());
			if (empty) return invalid(ctx, empty, 'Enter a value.');
			const n = Number(rank.value);
			if (
				spec.rank &&
				(!Number.isInteger(n) || n < 1 || n > (kind.endsWith('Percent') ? 100 : 1000))
			)
				return invalid(ctx, rank, 'Enter a whole number between {min} and {max}.', {
					min: 1,
					max: kind.endsWith('Percent') ? 100 : 1000,
				});
			const style =
				withSelect.value === 'custom' && custom
					? custom
					: (HIGHLIGHT_STYLES.find(([id]) => id === withSelect.value)?.[2] ??
						HIGHLIGHT_STYLES[0]![2]);
			const rule = quickRule(
				kind,
				inputs.map((i) => i.value),
				structuredClone(style),
				{
					rank: n,
					unique: dupSelect.value === 'unique',
					period: TIME_PERIOD_LABELS.find(([p]) => p === period.value)?.[0] ?? 'today',
				},
			);
			t.session.addConditionalFormat(t.sheet, { ranges: t.ranges, rules: [rule] });
			return rule;
		},
	});
}
