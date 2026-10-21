// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The fill-in fields of a form template, for a host that fills it without the editor: each
// legacy form field and content control with a stable id, its type, its constraints, its value
// and a label read from the layout around it (form-fill-labels.ts).
import { isContentControlLocked } from './content-control-model';
import { placeholderTextOf } from './content-control-edit';
import { formCheckboxChecked, formDropDownIndex } from './form-fields';
import { labelParts, labelText, numberedLabels } from './form-fill-labels';
import {
	formFieldSites,
	legacyResult,
	siteCell,
	type FieldSite,
	type FormFieldStory,
} from './form-fill-sites';
import type { DocumentModel, TextRun } from './model';
import { shownText } from './style-ref';

export type FormFieldType = 'text' | 'date' | 'number' | 'checkbox' | 'dropdown';
export type { FormFieldStory } from './form-fill-sites';

/** A drop-down or combo box entry: the text it shows and the value it stands for. */
export interface FormFieldChoice {
	display: string;
	value: string;
}

export interface FormFieldInfo {
	/**
	 * Stable for the same file: `field:<name>` for a legacy field whose name no other field has,
	 * `control:<id>` or `tag:<tag>` for a content control, else `field#<n>` or `control#<n>`,
	 * the field's place among the legacy fields or the controls in reading order.
	 */
	id: string;
	/** `legacy` for a Word 97 form field (`w:ffData`), `control` for a content control (`w:sdt`). */
	source: 'legacy' | 'control';
	type: FormFieldType;
	/** The label read from the layout, numbered when several fields have it. */
	label: string;
	/** What `label` was built from, the most general first (a row's label, then the field's own). */
	labelParts: string[];
	/** The text it holds, or whether a check box is checked; empty while it shows a placeholder. */
	value: string | boolean;
	/** A legacy text field's default text, or a legacy check box's default state. */
	defaultValue?: string | boolean;
	/** What an empty control shows. */
	placeholder?: string;
	/** A drop-down's or combo box's entries. */
	choices?: FormFieldChoice[];
	/** A combo box also takes text that is none of its entries. */
	allowsOtherText?: boolean;
	maxLength?: number;
	/** A date picture (`dd-MMM-yyyy`), a number picture (`#,##0.00`) or a text case (`UPPERCASE`). */
	format?: string;
	/** A date control's language (`en-US`). */
	locale?: string;
	/** Whether the text may hold line breaks. */
	multiLine?: boolean;
	/** A legacy field's `w:name`, as written (several fields can share one). */
	name?: string;
	/** A control's title (`w:alias`) and tag (`w:tag`). */
	alias?: string;
	tag?: string;
	/** A legacy field's help text. */
	help?: string;
	story: FormFieldStory;
	/** Whether `fillFormFields` sets it; when not, `reason` says why. */
	fillable: boolean;
	reason?: 'disabled' | 'calculated' | 'locked' | 'bound' | 'nested table';
}

/** A text field's result is empty when it is only Word's en-space placeholder or spaces. */
const resultValue = (text: string) => (/^\s*$/u.test(text) ? '' : text);

const contentText = (runs: readonly TextRun[]) =>
	runs
		.filter((run) => !run.fieldChar && run.fieldCode === undefined)
		.map(shownText)
		.join('');

/** The text a site's content shows: a legacy field's result, a control's runs or paragraphs. */
export function siteText(site: FieldSite): string {
	if (site.kind === 'legacy') {
		const result = legacyResult(site);
		return result ? contentText(site.paragraph.runs.slice(result.from, result.to + 1)) : '';
	}
	if (site.kind === 'inline') return contentText(site.paragraph.runs.slice(site.from, site.to + 1));
	const paragraphs =
		site.kind === 'cell'
			? (siteCell(site)?.paragraphs ?? [])
			: (site.paragraphs ?? []).map((index) => site.blocks[index]!);
	return paragraphs
		.map((item) => (item.type === 'paragraph' ? contentText(item.runs) : ''))
		.join('\n');
}

type Described = Omit<FormFieldInfo, 'id' | 'label' | 'labelParts'>;

function describeLegacy(site: FieldSite): Described {
	const field = site.formField!;
	const input = field.textInput;
	const base = {
		source: 'legacy' as const,
		story: site.story,
		...(field.name && { name: field.name }),
		...(field.helpText?.type === 'text' && field.helpText.value && { help: field.helpText.value }),
	};
	const blocked = (): Pick<FormFieldInfo, 'fillable' | 'reason'> => {
		if (site.nested) return { fillable: false, reason: 'nested table' };
		if (field.enabled === false) return { fillable: false, reason: 'disabled' };
		// A field inside a control that locks its content cannot change.
		const around = site.paragraph.runs[site.from]?.contentControl;
		if (isContentControlLocked(around, 'content')) return { fillable: false, reason: 'locked' };
		return { fillable: true };
	};
	if (field.kind === 'checkbox')
		return {
			...base,
			type: 'checkbox',
			value: formCheckboxChecked(field),
			defaultValue: field.checkBox?.default ?? false,
			...blocked(),
		};
	if (field.kind === 'dropDown') {
		const entries = field.dropDown?.entries ?? [];
		const index = formDropDownIndex(field);
		return {
			...base,
			type: 'dropdown',
			value: index === undefined ? '' : entries[index]!,
			choices: entries.map((entry) => ({ display: entry, value: entry })),
			...blocked(),
		};
	}
	const kind = input?.type;
	const computed = kind === 'currentDate' || kind === 'currentTime' || kind === 'calculated';
	return {
		...base,
		type: kind === 'date' ? 'date' : kind === 'number' ? 'number' : 'text',
		value: resultValue(siteText(site)),
		...(input?.default !== undefined && { defaultValue: input.default }),
		...(input?.maxLength && { maxLength: input.maxLength }),
		...(input?.format && { format: input.format }),
		...(computed ? { fillable: false, reason: 'calculated' as const } : blocked()),
	};
}

function describeControl(site: FieldSite): Described {
	const control = site.control!;
	const fillable: Pick<FormFieldInfo, 'fillable' | 'reason'> = site.nested
		? { fillable: false, reason: 'nested table' }
		: isContentControlLocked(control, 'content')
			? { fillable: false, reason: 'locked' }
			: control.dataBound
				? { fillable: false, reason: 'bound' }
				: { fillable: true };
	const base = {
		source: 'control' as const,
		story: site.story,
		...(control.alias && { alias: control.alias }),
		...(control.tag && { tag: control.tag }),
		...fillable,
	};
	if (control.kind === 'checkbox')
		return { ...base, type: 'checkbox', value: control.checked ?? false };
	const value = control.showingPlaceholder ? '' : siteText(site);
	const placeholder = { placeholder: placeholderTextOf(control) };
	if (control.kind === 'dropDownList' || control.kind === 'comboBox')
		return {
			...base,
			type: 'dropdown',
			value,
			choices: (control.items ?? []).map(({ display, value: item }) => ({ display, value: item })),
			...(control.kind === 'comboBox' && { allowsOtherText: true }),
			...placeholder,
		};
	if (control.kind === 'date')
		return {
			...base,
			type: 'date',
			value,
			...(control.dateFormat && { format: control.dateFormat }),
			...(control.dateLocale && { locale: control.dateLocale }),
			...placeholder,
		};
	return {
		...base,
		type: 'text',
		value,
		...((control.kind === 'richText' || control.multiLine) && { multiLine: true }),
		...placeholder,
	};
}

const TYPE_NAMES: Record<FormFieldType, string> = {
	text: 'Text',
	date: 'Date',
	number: 'Number',
	checkbox: 'Check box',
	dropdown: 'Drop-down',
};

/** The id of each site, from its name, control id or tag when no other field shares it. */
export function siteIds(sites: readonly FieldSite[]): string[] {
	const count = (keys: (string | undefined)[]) => {
		const counts = new Map<string, number>();
		for (const key of keys) if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
		return (key: string | undefined): key is string => Boolean(key) && counts.get(key!) === 1;
	};
	const legacy = sites.filter((site) => site.kind === 'legacy');
	const controls = sites.filter((site) => site.kind !== 'legacy');
	const uniqueName = count(legacy.map((site) => site.formField!.name));
	// A control without `w:id` gets a generated one, which another load would not repeat.
	const ownId = (site: FieldSite) =>
		site.control!.id.startsWith('sdt-gen-') ? undefined : site.control!.id;
	const uniqueId = count(controls.map(ownId));
	const uniqueTag = count(controls.map((site) => site.control!.tag));
	return sites.map((site) => {
		if (site.kind === 'legacy') {
			const name = site.formField!.name;
			return uniqueName(name) ? `field:${name}` : `field#${legacy.indexOf(site) + 1}`;
		}
		const id = ownId(site);
		const tag = site.control!.tag;
		if (uniqueId(id)) return `control:${id}`;
		if (uniqueTag(tag)) return `tag:${tag}`;
		return `control#${controls.indexOf(site) + 1}`;
	});
}

/** Each site described, with its id and label. */
export function describeSites(sites: readonly FieldSite[]): FormFieldInfo[] {
	const ids = siteIds(sites);
	const described = sites.map((site) =>
		site.kind === 'legacy' ? describeLegacy(site) : describeControl(site),
	);
	const parts = sites.map((site) => labelParts(site, sites));
	const labels = numberedLabels(
		parts.map((items, index) => {
			const info = described[index]!;
			const own = labelText(items.join(': '));
			return own || info.alias || info.tag || info.name || TYPE_NAMES[info.type];
		}),
	);
	return described.map((info, index) => ({
		id: ids[index]!,
		label: labels[index]!,
		labelParts: parts[index]!,
		...info,
	}));
}

/** The fill-in fields of a document in reading order: the body, then headers and footers. */
export const listFormFields = (model: DocumentModel): FormFieldInfo[] =>
	describeSites(formFieldSites(model));
