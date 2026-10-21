// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Filling a form without the editor: setting legacy form fields and content controls by id, as a
// person filling the form in Word would. A text field's result runs give way to one run in the
// first one's formatting; a check box or drop-down changes its data; a control leaves its
// placeholder (Word's Placeholder Text style goes with it) and a temporary control is removed,
// as the editor does. STYLEREF results that show the filled text are brought up to date.
import {
	afterEdit,
	checkboxState,
	placeholderTextOf,
	selectListItem,
	setControlDate,
	withValue,
	type ControlChange,
} from './content-control-edit';
import { selectFormDropDownEntry, setFormCheckbox } from './form-fields';
import { describeSites, siteText as siteTextOf, type FormFieldInfo } from './form-fill-fields';
import { formFieldSites, legacyResult, siteCell, type FieldSite } from './form-fill-sites';
import { prepareValue, type FormFieldValue, type PreparedValue } from './form-fill-values';
import type { DocumentModel, Paragraph, TextRun } from './model';
import { refreshStyleRefFields } from './style-ref-fields';

/** What happened to one value. */
export interface FormFillResult {
	id: string;
	/** Whether the field holds the value now (also when it held it already). */
	ok: boolean;
	/** Why the value was not set. */
	reason?: string;
	/** What the field shows now: the text written (a formatted date or number) or the state. */
	value?: string | boolean;
}

export interface FormFillOptions {
	/** Bring body STYLEREF results up to date as Update Fields does (default true). */
	updateStyleRefs?: boolean;
}

const ENSPACES = '\u2002'.repeat(5);
const PLACEHOLDER_STYLE = 'PlaceholderText';
/** Run properties that make a run something other than formatted text. */
const STRUCTURE = [
	'fieldChar',
	'fieldCode',
	'formField',
	'fieldInstanceId',
	'image',
	'equation',
	'noteReference',
	'noteMark',
	'noteSeparator',
	'break',
	'symbol',
	'positionalTab',
	'revision',
	'contentControl',
] as const;

/** A text run in the formatting of `source`, without what made `source` more than text. */
function textLike(source: TextRun | undefined, text: string): TextRun {
	const run: TextRun = { ...source, text };
	for (const key of STRUCTURE) delete run[key];
	if (run.style === PLACEHOLDER_STYLE) delete run.style;
	return run;
}

const isText = (run: TextRun) => !run.fieldChar && run.fieldCode === undefined;
const firstText = (runs: readonly TextRun[]) =>
	runs.find((run) => isText(run) && run.text) ?? runs.find(isText);

/**
 * Puts the run `make` builds in place of runs `from` to `to`. When a run of the paragraph keeps
 * properties the model does not hold, the writer must keep the paragraph's runs as they are
 * split: then the first text run of the range takes the new text and the others are emptied.
 */
function replaceRuns(
	paragraph: Paragraph,
	from: number,
	to: number,
	make: (source: TextRun | undefined) => TextRun,
	each?: (run: TextRun) => TextRun,
): void {
	const runs = paragraph.runs;
	const source = firstText(runs.slice(from, to + 1));
	const keep = source && runs.some((run) => run.sourceRunPropertiesXml !== undefined);
	if (!keep) {
		runs.splice(from, Math.max(0, to - from + 1), make(source));
		return;
	}
	for (let index = from; index <= to; index++) {
		const run = runs[index]!;
		if (run === source) runs[index] = make(source);
		else if (isText(run)) runs[index] = (each ?? ((item) => item))({ ...run, text: '' });
	}
}

/** The run a control's new content is written in. */
function controlRun(source: TextRun | undefined, change: ControlChange, placeholder: boolean) {
	const run = textLike(source, change.text);
	if (change.font) run.fontFamily = change.font;
	if (placeholder) run.style = PLACEHOLDER_STYLE;
	return run;
}

/** The control change a prepared value makes, or the reason it cannot. */
function controlChange(site: FieldSite, prepared: PreparedValue): ControlChange | string {
	const control = site.control!;
	if ('reason' in prepared) return prepared.reason;
	if ('clear' in prepared) {
		const { fullDate: _date, ...rest } = control;
		return { control: { ...rest, showingPlaceholder: true }, text: placeholderTextOf(control) };
	}
	if ('checked' in prepared) return afterEdit(checkboxState(control, prepared.checked));
	if ('choice' in prepared) return selectListItem(control, prepared.choice) ?? 'no such entry';
	if (prepared.date) return setControlDate(control, prepared.date);
	return afterEdit({ control: withValue(control), text: prepared.text });
}

/** Writes a control's new content where the control is. */
function writeControl(site: FieldSite, change: ControlChange): void {
	const placeholder = Boolean(change.control.showingPlaceholder && !change.removed);
	const control = change.removed ? undefined : change.control;
	// Every run of an inline control carries it, and they all carry its new state.
	const owned = (run: TextRun): TextRun => {
		const next = { ...run };
		if (control) next.contentControl = control;
		else delete next.contentControl;
		return next;
	};
	if (site.kind === 'inline') {
		replaceRuns(
			site.paragraph,
			site.from,
			site.to,
			(source) => owned(controlRun(source, change, placeholder)),
			owned,
		);
		return;
	}
	const paragraphs =
		site.kind === 'cell'
			? siteCell(site)!.paragraphs
			: site.paragraphs!.map((index) => site.blocks[index] as Paragraph);
	// The control keeps its first paragraph, holding the new text in the first text's formatting.
	const formatting = firstText(paragraphs.flatMap((item) => item.runs));
	const first: Paragraph = { ...paragraphs[0]!, runs: [...paragraphs[0]!.runs] };
	replaceRuns(first, 0, first.runs.length - 1, (source) =>
		controlRun(source ?? formatting, change, placeholder),
	);
	if (site.kind === 'cell') {
		const cell = siteCell(site)!;
		cell.paragraphs = [first];
		if (control) cell.contentControl = control;
		else delete cell.contentControl;
		return;
	}
	if (control) first.contentControl = control;
	else delete first.contentControl;
	site.blocks.splice(site.paragraphs![0]!, site.paragraphs!.length, first);
}

/** Sets a legacy field from a prepared value; the reason when it cannot. */
function writeLegacy(site: FieldSite, prepared: PreparedValue): string | undefined {
	if ('reason' in prepared) return prepared.reason;
	const runs = site.paragraph.runs;
	const begin = runs[site.from]!;
	const field = begin.formField!;
	if ('checked' in prepared) {
		runs[site.from] = { ...begin, formField: setFormCheckbox(field, prepared.checked) };
		return undefined;
	}
	if ('choice' in prepared) {
		const picked = selectFormDropDownEntry(field, prepared.choice);
		if (!picked) return 'no such entry';
		runs[site.from] = { ...begin, formField: picked };
		return undefined;
	}
	const result = legacyResult(site);
	if (!result) return 'the field has no result to hold text';
	const text = 'clear' in prepared ? ENSPACES : prepared.text;
	replaceRuns(site.paragraph, result.from, result.to, (source) =>
		textLike(source?.field ? source : { ...begin, field: { instr: 'FORMTEXT' } }, text),
	);
	return undefined;
}

/** What the field holds after the change, as the listing reports it. */
const shownValue = (prepared: PreparedValue, info: FormFieldInfo, change?: ControlChange) =>
	'checked' in prepared
		? prepared.checked
		: 'clear' in prepared
			? ''
			: 'choice' in prepared
				? (info.choices?.[prepared.choice]?.display ?? '')
				: (change?.text ?? ('text' in prepared ? prepared.text : ''));

/** Whether a legacy field already holds what the value would give it. */
function legacyHolds(info: FormFieldInfo, prepared: PreparedValue): boolean {
	if ('checked' in prepared) return info.value === prepared.checked;
	if ('clear' in prepared) return info.value === '';
	if ('choice' in prepared) return info.value === info.choices?.[prepared.choice]?.display;
	return 'text' in prepared && info.value === prepared.text;
}

/** JSON with every object's keys in order, so that objects with the same fields compare equal. */
const canonical = (value: unknown): string =>
	JSON.stringify(value, (_key, item: unknown) =>
		item && typeof item === 'object' && !Array.isArray(item)
			? Object.fromEntries(Object.entries(item).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
			: item,
	);

/** Whether a control already shows what the change would make it show: no edit, nothing to do. */
const controlHolds = (site: FieldSite, change: ControlChange): boolean =>
	canonical(change.control) === canonical(site.control) && change.text === siteTextOf(site);

/**
 * Sets fields by id (as `listFormFields` gives them) and returns the filled model, which
 * `loadDocx(...).save(model)` writes back over the source package. `model` itself is not
 * changed. A value that does not fit its field is not set, and its result says why.
 */
export function fillFormFields(
	model: DocumentModel,
	values: Readonly<Record<string, FormFieldValue>>,
	options: FormFillOptions = {},
): { model: DocumentModel; results: FormFillResult[] } {
	const next = structuredClone(model);
	const sites = formFieldSites(next);
	const infos = describeSites(sites);
	const results: FormFillResult[] = [];
	const edits: { at: number; result: FormFillResult; apply: () => string | undefined }[] = [];
	for (const [id, value] of Object.entries(values)) {
		const at = infos.findIndex((info) => info.id === id);
		const info = infos[at];
		const site = sites[at];
		const refuse = (reason: string) => results.push({ id, ok: false, reason });
		if (!info || !site) {
			refuse('no field has this id');
			continue;
		}
		if (!info.fillable) {
			refuse(info.reason ?? 'the field cannot be filled');
			continue;
		}
		const prepared = prepareValue(info, value);
		if ('reason' in prepared) {
			refuse(prepared.reason);
			continue;
		}
		const change = site.kind === 'legacy' ? undefined : controlChange(site, prepared);
		if (typeof change === 'string') {
			refuse(change);
			continue;
		}
		if (change ? controlHolds(site, change) : legacyHolds(info, prepared)) {
			results.push({ id, ok: true, value: info.value });
			continue;
		}
		const result: FormFillResult = { id, ok: true, value: shownValue(prepared, info, change) };
		results.push(result);
		edits.push({
			at,
			result,
			apply: () => (change ? (writeControl(site, change), undefined) : writeLegacy(site, prepared)),
		});
	}
	// Later fields first, so that an edit never moves the runs or blocks of one still to come.
	for (const edit of edits.sort((a, b) => b.at - a.at)) {
		const reason = edit.apply();
		if (!reason) continue;
		edit.result.ok = false;
		edit.result.reason = reason;
		delete edit.result.value;
	}
	const filled =
		options.updateStyleRefs === false ? next : refreshStyleRefFields(next, { body: true });
	return { model: filled, results };
}
