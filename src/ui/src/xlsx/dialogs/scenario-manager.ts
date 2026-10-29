// Data > What-If Analysis > Scenario Manager: add, edit, delete and show the sheet's scenarios.
// Merge copies the scenarios of another sheet; Summary builds the report on a new sheet.
import { formatRange, previewScenario, scenariosOf, type Scenario } from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, el, field, listBox, text, textArea, textInput } from './fields';
import { button, showDialog } from './frame';
import { openScenarioMerge, openScenarioSummary } from './scenario-extras';

type Target = NonNullable<ReturnType<typeof target>>;

/** Add or edit: the name, changing cells and comment, then the values of every changing cell. */
async function editScenario(
	ctx: EditorContext,
	t: Target,
	existing?: { scenario: Scenario; index: number },
): Promise<boolean> {
	const name = textInput(ctx, existing?.scenario.name ?? '');
	const cells = textInput(
		ctx,
		existing
			? existing.scenario.cells.map((c) => c.ref).join(',')
			: t.ranges.map((r) => formatRange(r)).join(','),
	);
	const comment = textArea(ctx, existing?.scenario.comment ?? '', 3);
	comment.setAttribute('aria-label', ctx.t('Comment:'));
	const locked = checkbox(ctx, 'Prevent changes', existing?.scenario.locked ?? true);
	const hidden = checkbox(ctx, 'Hide', existing?.scenario.hidden ?? false);
	const draft = await showDialog<{ name: string; cells: string; comment: string }>(ctx, {
		name: 'scenario-edit',
		heading: existing ? 'Edit Scenario' : 'Add Scenario',
		okLabel: 'OK',
		body: [
			field(ctx, 'Scenario name:', name),
			field(ctx, 'Changing cells:', cells),
			(() => {
				const label = el(ctx, 'label', 'xve-field');
				const span = el(ctx, 'span', 'xve-field-label');
				span.textContent = ctx.t('Comment:');
				label.append(span, comment);
				return label;
			})(),
			locked.wrapper,
			hidden.wrapper,
		],
		opened: () => name.focus(),
		submit: () => ({ name: name.value, cells: cells.value, comment: comment.value }),
	});
	if (!draft) return false;
	// Validate the name and cells before asking for values.
	const probe = { ...draft, locked: locked.input.checked, hidden: hidden.input.checked };
	const addresses = previewScenario(t.workbook, t.sheet, probe, existing?.index);
	const inputs = new Map<string, HTMLInputElement>();
	const body = addresses.map(({ ref, value }) => {
		const input = textInput(ctx, value);
		inputs.set(ref, input);
		return field(ctx, `${ref}:`, input);
	});
	const values = await showDialog<Record<string, string>>(ctx, {
		name: 'scenario-values',
		heading: 'Scenario Values',
		body: [text(ctx, 'Enter values for each of the changing cells.'), ...body],
		opened: () => inputs.values().next().value?.focus(),
		submit: () => Object.fromEntries([...inputs].map(([ref, input]) => [ref, input.value])),
	});
	if (!values) return false;
	const full = { ...probe, values };
	if (existing) t.session.editScenario(t.sheet, existing.index, full);
	else t.session.addScenario(t.sheet, full);
	return true;
}

export function openScenarioManager(ctx: EditorContext): Promise<undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const list = listBox(
		ctx,
		'Scenarios:',
		() => describe(),
		() => void run('edit'),
	);
	const details = el(ctx, 'div', 'xve-note');
	details.setAttribute('aria-live', 'polite');
	const add = button(ctx, 'Add...');
	const edit = button(ctx, 'Edit...');
	const del = button(ctx, 'Delete');
	const merge = button(ctx, 'Merge...');
	const summary = button(ctx, 'Summary...');
	const scenarios = (): Scenario[] => scenariosOf(t.ws);
	const chosen = (): number | undefined => {
		const value = list.value();
		return value === undefined ? undefined : Number(value);
	};
	function describe(): void {
		const index = chosen();
		const scenario = index === undefined ? undefined : scenarios()[index];
		edit.disabled = del.disabled = !scenario;
		summary.disabled = !scenarios().length;
		if (!scenario) {
			details.textContent = scenarios().length
				? ctx.t('Select a scenario to see its changing cells.')
				: ctx.t('No Scenarios defined. Choose Add to add scenarios.');
			return;
		}
		const parts = [
			`${ctx.t('Changing cells:')} ${scenario.cells.map((c) => c.ref).join(', ')}`,
			...(scenario.comment ? [`${ctx.t('Comment:')} ${scenario.comment}`] : []),
		];
		details.textContent = parts.join('\n');
		details.style.whiteSpace = 'pre-line';
	}
	const refresh = (): void => {
		list.setItems(scenarios().map((s, i) => [String(i), s.name] as const));
		describe();
	};
	const run = async (action: 'add' | 'edit' | 'delete'): Promise<void> => {
		const index = chosen();
		try {
			if (action === 'add') await editScenario(ctx, t, undefined);
			else if (index !== undefined) {
				const scenario = scenarios()[index];
				if (!scenario) return;
				if (action === 'edit') await editScenario(ctx, t, { scenario, index });
				else t.session.deleteScenario(t.sheet, index);
			}
		} catch (error) {
			ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'error');
		}
		refresh();
	};
	add.addEventListener('click', () => void run('add'));
	edit.addEventListener('click', () => void run('edit'));
	del.addEventListener('click', () => void run('delete'));
	merge.addEventListener('click', () => void openScenarioMerge(ctx, t).then(refresh));
	summary.addEventListener('click', () => void openScenarioSummary(ctx, t).then(refresh));
	refresh();
	const first = scenarios()[t.ws.scenarios?.show ?? 0] ? (t.ws.scenarios?.show ?? 0) : 0;
	if (scenarios().length) list.select(String(first));
	return showDialog<undefined>(ctx, {
		name: 'scenario-manager',
		heading: 'Scenario Manager',
		okLabel: 'Show',
		cancelLabel: 'Close',
		buttons: [add, del, edit, merge, summary],
		body: [list.element, details],
		opened: () => list.element.focus(),
		// Show keeps the dialog open, as Excel does.
		submit: () => {
			const index = chosen();
			if (index === undefined) {
				ctx.toast(ctx.t('Select a scenario to show.'), 'warning');
				return undefined;
			}
			t.session.showScenario(t.sheet, index);
			return undefined;
		},
	}).then(() => undefined);
}
