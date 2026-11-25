import { DESIGN_COLOR_SLOTS, designColors, designSpacing, designTheme } from 'ooxml-core/docx';
import type { DesignEdit, DesignSpacing } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { dialogButton, labelled, numberInput, selectOf } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import { localeOf, translate, type LocalizationKey } from './localization';
import { canEditDesign } from './design-ribbon-controls';

export const colorLabels = {
	lt1: 'design.light1',
	dk1: 'design.dark1',
	lt2: 'design.light2',
	dk2: 'design.dark2',
	accent1: 'design.accent1',
	accent2: 'design.accent2',
	accent3: 'design.accent3',
	accent4: 'design.accent4',
	accent5: 'design.accent5',
	accent6: 'design.accent6',
	hlink: 'design.hyperlink',
	folHlink: 'design.followedHyperlink',
} satisfies Record<(typeof DESIGN_COLOR_SLOTS)[number], LocalizationKey>;

type CustomMenu = 'colors' | 'fonts' | 'spacing';
export const customTitles = {
	colors: 'design.customColors',
	fonts: 'design.customFonts',
	spacing: 'design.customSpacing',
} satisfies Record<CustomMenu, LocalizationKey>;

type EditFields = { fields: HTMLElement[]; read(): DesignEdit | undefined };

function textInput(value: string): HTMLInputElement {
	const input = document.createElement('input');
	input.type = 'text';
	input.maxLength = 100;
	input.value = value.slice(0, 100);
	return input;
}

function namedFields(core: EditorCore, kind: 'colors' | 'fonts'): EditFields {
	const theme = designTheme(core.model.theme);
	const name = textInput(
		(kind === 'fonts' ? theme.fontSchemeName : theme.colorSchemeName) ??
			translate(core.locale, 'Custom'),
	);
	const fields: HTMLElement[] = [labelled('design.schemeName', name)];
	if (kind === 'fonts') {
		const major = textInput(theme.fonts.major.latin ?? 'Calibri Light');
		const minor = textInput(theme.fonts.minor.latin ?? 'Calibri');
		fields.push(labelled('design.headingFont', major), labelled('design.bodyFont', minor));
		return {
			fields,
			read: () => {
				const values = [name, major, minor].map((input) => input.value.trim());
				const [nameValue, majorFont, minorFont] = values;
				if (nameValue && majorFont && minorFont && values.every((value) => value.length <= 100))
					return { kind, name: nameValue, majorFont, minorFont };
			},
		};
	}
	const colors = { ...designColors(), ...theme.colors };
	const controls = DESIGN_COLOR_SLOTS.map((slot) => {
		const input = document.createElement('input');
		input.type = 'color';
		input.value = `#${colors[slot]?.replace(/^#/, '') ?? '000000'}`;
		fields.push(labelled(colorLabels[slot], input));
		return { slot, input };
	});
	return {
		fields,
		read: () => {
			const value = name.value.trim();
			if (value && value.length <= 100)
				return {
					kind,
					name: value,
					colors: Object.fromEntries(
						controls.map(({ slot, input }) => [slot, input.value.slice(1).toUpperCase()]),
					),
				};
		},
	};
}

function spacingFields(spacing: DesignSpacing): EditFields {
	const before = numberInput(0, 1584, 0.05);
	const after = numberInput(0, 1584, 0.05);
	const rule = selectOf([
		['auto', 'Multiple'],
		['exact', 'Exactly'],
		['atLeast', 'At least'],
	]);
	const line = numberInput(0.01, 132, 0.01);
	const label = labelled('design.lines', line);
	let currentRule = spacing.rule;
	const updateUnits = () => {
		const auto = rule.value === 'auto';
		const key = auto ? 'design.lines' : 'design.linePoints';
		label.firstElementChild?.replaceChildren(key);
		line.dataset.localearialabel = key;
		line.max = auto ? '132' : '1584';
		line.min = auto ? '0.01' : '0.05';
		line.step = auto ? 'any' : '0.05';
	};
	before.value = String(spacing.before / 20);
	after.value = String(spacing.after / 20);
	rule.value = spacing.rule;
	line.value = String(spacing.line / (spacing.rule === 'auto' ? 240 : 20));
	updateUnits();
	rule.addEventListener('change', () => {
		if (currentRule === 'auto' && rule.value !== 'auto') line.value = '12';
		if (currentRule !== 'auto' && rule.value === 'auto') line.value = '1';
		currentRule = rule.value === 'exact' || rule.value === 'atLeast' ? rule.value : 'auto';
		updateUnits();
		const dialog = rule.closest<HTMLElement>('office-ui-dialog');
		if (dialog && dialog.dataset.editorLocale) localizeDialog(dialog, localeOf(dialog));
	});
	return {
		fields: [
			labelled('design.beforePoints', before),
			labelled('design.afterPoints', after),
			labelled('Line spacing', rule),
			label,
		],
		read: () => {
			if (![before, after, line].every((input) => input.value !== '' && input.validity.valid))
				return;
			return {
				kind: 'spacing',
				spacing: {
					before: Math.round(before.valueAsNumber * 20),
					after: Math.round(after.valueAsNumber * 20),
					line: Math.round(line.valueAsNumber * (currentRule === 'auto' ? 240 : 20)),
					rule: currentRule,
				},
			};
		},
	};
}

export function openDesignDialog(
	core: EditorCore,
	anchor: HTMLButtonElement,
	kind: CustomMenu | 'default',
	apply: (edit?: DesignEdit) => void,
): void {
	const root = anchor.getRootNode();
	if (!(root instanceof ShadowRoot)) return;
	const element = createDialogShell(
		kind === 'default' ? 'Set as Default' : customTitles[kind],
		'dve-format-dialog dve-design-dialog',
	);
	element.dataset.editorLocale = core.locale;
	const cancel = dialogButton('Cancel');
	const ok = dialogButton(kind === 'default' ? 'OK' : 'Save', true);
	const close = () => {
		setDialogOpen(element, false);
		element.remove();
		anchor.focus();
	};
	const form =
		kind === 'default'
			? undefined
			: kind === 'spacing'
				? spacingFields(designSpacing(core.model))
				: namedFields(core, kind);
	if (form) element.append(...form.fields);
	else {
		const message = document.createElement('p');
		message.textContent = 'design.confirmDefault';
		element.append(message);
	}
	const refresh = () => {
		ok.disabled = !canEditDesign(core) || Boolean(form && !form.read());
	};
	element.addEventListener('input', refresh);
	element.addEventListener('change', refresh);
	cancel.addEventListener('click', close);
	ok.addEventListener('click', () => {
		const value = form?.read();
		if (!canEditDesign(core) || (form && !value)) return;
		close();
		apply(value);
	});
	onDialogDismiss(element, close);
	element.append(dialogActions(cancel, ok));
	localizeDialog(element, core.locale);
	root.append(element);
	refresh();
	setDialogOpen(element, true);
}
