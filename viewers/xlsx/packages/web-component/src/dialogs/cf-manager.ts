// Conditional Formatting Rules Manager: every rule of the sheet (or those touching the selection)
// in priority order, with Applies to and Stop If True. Edits are staged and written with one
// undo step on OK or Apply.
import { type ConditionalFormat, rangesIntersect } from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { describeRule, formatRefs, hasStopIfTrue, parseRefs, previewBox } from './cf-common.js';
import { type RuleEdit, editRuleDialog } from './cf-rule.js';
import { el, field, invalid, row, select } from './fields.js';
import { button, showDialog } from './frame.js';

/** Flattens a sheet's formats into one entry per rule, sorted by priority. */
export function stageRules(formats: ConditionalFormat[]): RuleEdit[] {
	return formats
		.flatMap((cf) =>
			cf.rules.map((rule) => ({ rule: structuredClone(rule), ranges: structuredClone(cf.ranges) })),
		)
		.sort((a, b) => a.rule.priority - b.rule.priority);
}

/** One conditional format per staged rule, priorities 1..n in list order. */
export function unstageRules(entries: RuleEdit[]): ConditionalFormat[] {
	return entries.map((entry, i) => ({
		ranges: entry.ranges,
		rules: [{ ...entry.rule, priority: i + 1 }],
	}));
}

export function cfManagerDialog(ctx: EditorContext): Promise<ConditionalFormat[] | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	let entries = stageRules(t.ws.conditionalFormats);
	let selected = entries.length ? 0 : -1;
	const scope = select(
		ctx,
		[
			['selection', 'Current Selection'],
			['sheet', 'This Worksheet'],
		],
		'selection',
	);
	const table = el(ctx, 'table');
	table.setAttribute('role', 'grid');
	table.setAttribute('aria-label', ctx.t('Rules'));
	table.tabIndex = 0;
	const wrap = el(ctx, 'div', 'xve-results');
	wrap.append(table);
	const appliesInputs: HTMLInputElement[] = [];
	const visible = (): number[] =>
		entries.flatMap((e, i) =>
			scope.value === 'sheet' || e.ranges.some((r) => t.ranges.some((s) => rangesIntersect(r, s)))
				? [i]
				: [],
		);
	const render = (): void => {
		appliesInputs.length = 0;
		const head = el(ctx, 'tr');
		for (const key of ['Rule (applied in order shown)', 'Format', 'Applies to', 'Stop If True']) {
			const th = el(ctx, 'th');
			th.textContent = ctx.t(key);
			head.append(th);
		}
		const rows = visible().map((i) => {
			const entry = entries[i]!;
			const tr = el(ctx, 'tr');
			tr.dataset.index = String(i);
			tr.setAttribute('aria-selected', String(i === selected));
			tr.addEventListener('click', (event) => {
				if (selected === i || (event.target as HTMLElement).tagName === 'INPUT') {
					selected = i;
					for (const other of table.querySelectorAll('tr[data-index]'))
						other.setAttribute('aria-selected', String(other === tr));
					return;
				}
				selected = i;
				render();
			});
			const desc = el(ctx, 'td');
			desc.textContent = describeRule(ctx, entry.rule);
			const fmt = el(ctx, 'td');
			const preview = previewBox(ctx);
			preview.show(entry.rule);
			fmt.append(preview.element);
			const applies = el(ctx, 'td');
			const input = el(ctx, 'input', 'xve-input');
			input.value = formatRefs(entry.ranges);
			input.setAttribute('aria-label', ctx.t('Applies to'));
			input.dataset.index = String(i);
			input.addEventListener('change', () => {
				const ranges = parseRefs(input.value);
				if (ranges) entry.ranges = ranges;
			});
			appliesInputs.push(input);
			applies.append(input);
			const stop = el(ctx, 'td');
			const box = el(ctx, 'input');
			box.type = 'checkbox';
			box.setAttribute('aria-label', ctx.t('Stop If True'));
			box.disabled = !hasStopIfTrue(entry.rule);
			box.checked = hasStopIfTrue(entry.rule) && !!entry.rule.stopIfTrue;
			box.addEventListener('change', () => {
				if (!hasStopIfTrue(entry.rule)) return;
				if (box.checked) entry.rule.stopIfTrue = true;
				else delete entry.rule.stopIfTrue;
			});
			stop.append(box);
			tr.append(desc, fmt, applies, stop);
			return tr;
		});
		table.replaceChildren(head, ...rows);
	};
	/** Reads the Applies to fields into the entries; false (with a warning) when one is invalid. */
	const commitRanges = (): boolean => {
		for (const input of appliesInputs) {
			const entry = entries[Number(input.dataset.index)];
			const ranges = parseRefs(input.value);
			if (!ranges) {
				invalid(ctx, input, 'The reference is not valid.');
				return false;
			}
			if (entry) entry.ranges = ranges;
		}
		return true;
	};
	const move = (delta: number): void => {
		if (!commitRanges()) return;
		const to = selected + delta;
		if (selected < 0 || to < 0 || to >= entries.length) return;
		const [entry] = entries.splice(selected, 1);
		if (entry) entries.splice(to, 0, entry);
		selected = to;
		render();
	};
	const newRule = button(ctx, 'New Rule...');
	const editRule = button(ctx, 'Edit Rule...');
	const deleteRule = button(ctx, 'Delete Rule');
	const duplicate = button(ctx, 'Duplicate Rule');
	const up = button(ctx, 'Move Up');
	const down = button(ctx, 'Move Down');
	newRule.addEventListener('click', async () => {
		if (!commitRanges()) return;
		const result = await editRuleDialog(ctx);
		if (!result) return;
		entries.unshift(result);
		selected = 0;
		render();
	});
	editRule.addEventListener('click', async () => {
		if (!commitRanges()) return;
		const entry = entries[selected];
		if (!entry) return;
		const result = await editRuleDialog(ctx, entry);
		if (result) entries[selected] = result;
		render();
	});
	deleteRule.addEventListener('click', () => {
		if (selected < 0) return;
		entries.splice(selected, 1);
		selected = Math.min(selected, entries.length - 1);
		render();
	});
	duplicate.addEventListener('click', () => {
		if (!commitRanges()) return;
		const entry = entries[selected];
		if (!entry) return;
		entries.splice(selected, 0, structuredClone(entry));
		render();
	});
	up.addEventListener('click', () => move(-1));
	down.addEventListener('click', () => move(1));
	table.addEventListener('keydown', (event) => {
		if (event.target !== table) return;
		const list = visible();
		const at = list.indexOf(selected);
		const next =
			event.key === 'ArrowDown' ? list[at + 1] : event.key === 'ArrowUp' ? list[at - 1] : undefined;
		if (next === undefined) return;
		event.preventDefault();
		selected = next;
		render();
	});
	scope.addEventListener('change', render);
	const apply = (): boolean => {
		if (!commitRanges()) return false;
		const formats = unstageRules(entries);
		const current = unstageRules(stageRules(t.ws.conditionalFormats));
		if (JSON.stringify(formats) !== JSON.stringify(current))
			t.session.batch('Conditional Formatting Rules', () => {
				t.session.clearConditionalFormats(t.sheet);
				for (const format of formats) t.session.addConditionalFormat(t.sheet, format);
				// An added format goes first; replacing each in order restores the staged priorities
				// (the core renumbers by priority, ties in sheet order).
				formats.forEach((format, i) => t.session.replaceConditionalFormat(t.sheet, i, format));
			});
		entries = stageRules(formats);
		render();
		return true;
	};
	const applyButton = button(ctx, 'Apply');
	applyButton.addEventListener('click', () => void apply());
	render();
	return showDialog<ConditionalFormat[]>(ctx, {
		name: 'cf-manager',
		heading: 'Conditional Formatting Rules Manager',
		wide: true,
		body: [
			field(ctx, 'Show formatting rules for:', scope),
			row(ctx, newRule, editRule, deleteRule, duplicate, up, down),
			wrap,
		],
		buttons: [applyButton],
		opened: () => scope.focus(),
		submit: () => (apply() ? unstageRules(entries) : undefined),
	});
}
