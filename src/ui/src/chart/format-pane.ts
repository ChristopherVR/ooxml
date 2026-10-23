import { html, nothing, type PropertyValues, type TemplateResult } from 'lit';
import { live } from 'lit/directives/live.js';
import { repeat } from 'lit/directives/repeat.js';
import { html as staticHtml, unsafeStatic } from 'lit/static-html.js';
import type { ChartEditIntent, ChartSpace } from 'ooxml-core/chart';
import type { DrawingColorTheme } from 'ooxml-core/drawingml';

import { OfficeElement, controlStyles, flag } from '../base';
import rangeCss from '../form/range-control.css?raw';
import { defineCheckbox } from '../form/checkbox';
import { defineRadio } from '../form/radio';
import { definer, emit } from '../registry';
import css from './format-pane.css?raw';
import { translateChartFormatText, type ChartFormatTranslate } from './format-pane-common';
import {
	fieldTemplate,
	type ChartFormatControlTags,
	type FieldContext,
} from './format-pane-fields';
import { chartFormatFieldIntent } from './format-pane-intent';
import { buildChartFormatModel } from './format-pane-model';
import { chartSpaceFormatSource } from './format-pane-space';
import type {
	ChartFormatFieldValue,
	ChartFormatPaneMode,
	ChartFormatPaneModel,
	ChartFormatSection,
	ChartFormatTabId,
	ChartFormatText,
} from './format-pane-types';
import type { ChartFormatSource } from './format-pane-view';

/** What `office-chart-format-edit` carries: the field change and, where one exists, its intent. */
export interface ChartFormatEditDetail {
	/** The target key (`chartArea`, `series:0`, `axis:1`...). */
	target: string;
	/** The field id (`fill.color`, `series.gapWidth`...). */
	field: string;
	value: ChartFormatFieldValue;
	/** The neutral chart edit, or `null` for a field no neutral operation covers. */
	intent: ChartEditIntent | null;
}
export type ChartFormatEditEvent = CustomEvent<ChartFormatEditDetail>;
export type ChartFormatTargetEvent = CustomEvent<{ target: string }>;

interface Statics {
	editEvent: string;
	targetEvent: string;
	closeEvent: string;
	controlTags: ChartFormatControlTags;
}

const echo: ChartFormatTranslate = (key) => key;

/**
 * `<office-ui-chart-format-pane>`: the chart Format task pane every product shares (Format Chart
 * Area, Format Data Series, Format Axis, ...): the element picker, Chart Options and Text Options,
 * the icon tabs and their collapsible sections. Give it a `source` (what the chart reads as, see
 * `ChartFormatSource`) or a `chartSpace` (read through `chartSpaceFormatSource`) and a `target`
 * key; it applies nothing itself. Each committed field fires `office-chart-format-edit` with the
 * field, its value and the typed `ChartEditIntent` (`applyChartEditIntent` runs it over a
 * `ChartSpace`; other products map it onto their own edits), then repaints from the source the
 * host sets back. The picker fires `office-chart-format-target`, Close and Escape
 * `office-chart-format-close`. A source with a `slots.fill` name leaves that section to the host's
 * light-DOM child with the same `slot`. `translateLabel(key, params, fallback)` translates captions
 * (`chartFormat.*` keys with English fallbacks); `disabled` makes every control read-only. It lays
 * out in flow; a product docks it.
 */
export class OfficeUiChartFormatPane extends OfficeElement {
	static editEvent = 'office-chart-format-edit';
	static targetEvent = 'office-chart-format-target';
	static closeEvent = 'office-chart-format-close';
	static controlTags: ChartFormatControlTags = {
		select: 'select',
		radio: 'office-ui-radio',
		checkbox: 'office-ui-checkbox',
	};
	static override styles = controlStyles(rangeCss, css);
	static override properties = { disabled: flag };
	/** Properties a host may set before the element is defined; replayed through their setters. */
	static replayed: readonly string[] = [
		'source',
		'chartSpace',
		'theme',
		'target',
		'translateLabel',
	];
	/** Every control is read-only and nothing is emitted. */
	declare disabled: boolean;

	private sourceValue: ChartFormatSource | undefined;
	private spaceValue: ChartSpace | undefined;
	private themeValue: DrawingColorTheme | undefined;
	private targetValue = 'chartArea';
	private translator: ChartFormatTranslate = echo;
	private paneTab: ChartFormatTabId = 'fillLine';
	private paneMode: ChartFormatPaneMode = 'chart';
	private collapsed = new Set<string>();
	/** After an edit, write the model's values even into the focused box (a refused value). */
	private forceSync = false;
	/** The pane title, the host's accessible name. */
	private hostLabel = '';
	/** The value each box and select was last given, to tell typing from a stale value. */
	private synced = new WeakMap<Element, string>();

	constructor() {
		super();
		this.disabled = false;
	}

	override connectedCallback(): void {
		const self = this as unknown as Record<string, unknown>;
		for (const name of (this.constructor as typeof OfficeUiChartFormatPane).replayed) {
			if (!Object.prototype.hasOwnProperty.call(this, name)) continue;
			const value = self[name];
			delete self[name];
			self[name] = value;
		}
		super.connectedCallback();
	}

	/** What the pane reads; replaces `chartSpace`. */
	get source(): ChartFormatSource | undefined {
		return this.sourceValue;
	}
	set source(value: ChartFormatSource | undefined) {
		this.sourceValue = value;
		this.spaceValue = undefined;
		this.requestUpdate('source');
	}
	/** A chart part model, read through `chartSpaceFormatSource` with `theme`. */
	get chartSpace(): ChartSpace | undefined {
		return this.spaceValue;
	}
	set chartSpace(value: ChartSpace | undefined) {
		this.spaceValue = value;
		this.sourceValue = value && chartSpaceFormatSource(value, this.themeOptions());
		this.requestUpdate('chartSpace');
	}
	/** Resolves scheme colours of a `chartSpace`. */
	get theme(): DrawingColorTheme | undefined {
		return this.themeValue;
	}
	set theme(value: DrawingColorTheme | undefined) {
		this.themeValue = value;
		if (this.spaceValue) this.chartSpace = this.spaceValue;
	}
	/** The element the pane formats (`chartArea`, `series:0`, `axis:1`, ...). */
	get target(): string {
		return this.targetValue;
	}
	set target(value: string) {
		this.targetValue = value || 'chartArea';
		this.requestUpdate('target');
	}
	/** Translates captions: `(key, params, fallback)`, returning the key when it has none. */
	get translateLabel(): ChartFormatTranslate {
		return this.translator;
	}
	set translateLabel(value: ChartFormatTranslate | undefined) {
		this.translator = value ?? echo;
		this.requestUpdate('translateLabel');
	}

	private themeOptions() {
		return this.themeValue ? { theme: this.themeValue } : {};
	}

	private statics(): Statics {
		return this.constructor as unknown as Statics;
	}

	protected tr = (text: ChartFormatText): string =>
		translateChartFormatText(text, (key, params, fallback) =>
			this.translator(key, params, fallback),
		);

	private label(key: string, fallback: string): string {
		return this.tr({ key: `chartFormat.${key}`, fallback });
	}

	/** The pane for the current source, target, tab and mode, or undefined without a source. */
	protected model(): ChartFormatPaneModel | undefined {
		const source = this.sourceValue;
		return source && buildChartFormatModel(source, this.targetValue, this.paneTab, this.paneMode);
	}

	/** A committed field: its edit (with the neutral intent) goes to {@link commit}. */
	private edit(field: string, value: ChartFormatFieldValue): void {
		const source = this.sourceValue;
		const model = this.model();
		if (this.disabled || !source || !model) {
			this.refreshValues();
			return;
		}
		const view = source.element(model.target);
		this.commit({
			target: model.targetKey,
			field,
			value,
			intent: chartFormatFieldIntent(view, model.target, field, value),
		});
		this.refreshValues();
	}

	/**
	 * Fires the edit event. A product subclass that applies edits itself overrides this (PowerPoint
	 * maps the edit onto its own chart model); either way the pane repaints afterwards.
	 */
	protected commit(edit: ChartFormatEditDetail): void {
		emit<ChartFormatEditDetail>(this, this.statics().editEvent, edit);
	}

	/** Repaints every value from the model, even the box being typed in (a refused value). */
	protected refreshValues(): void {
		this.forceSync = true;
		this.requestUpdate();
	}

	protected pickTarget(target: string): void {
		this.targetValue = target;
		this.requestUpdate('target');
		emit(this, this.statics().targetEvent, { target });
	}

	protected requestClose(): void {
		emit(this, this.statics().closeEvent, undefined);
	}

	private toggleText(): void {
		this.paneMode = this.paneMode === 'text' ? 'chart' : 'text';
		this.paneTab = this.paneMode === 'text' ? 'textFill' : 'fillLine';
		this.requestUpdate();
	}

	private onKey = (event: KeyboardEvent): void => {
		event.stopPropagation();
		if (event.key === 'Escape' && !event.defaultPrevented) {
			event.preventDefault();
			this.requestClose();
		}
	};

	private preview = (id: string, value: number, text: string): void => {
		const field = CSS.escape(id);
		const box = this.renderRoot.querySelector<HTMLInputElement>(
			`[data-chart-format-field="${field}"] input`,
		);
		if (box) box.value = String(value);
		this.renderRoot
			.querySelector(`[data-chart-format-slider="${field}"] input`)
			?.setAttribute('aria-valuetext', text);
	};

	private section(section: ChartFormatSection, ctx: FieldContext): TemplateResult {
		const open = !this.collapsed.has(section.id);
		const toggle = () => {
			if (open) this.collapsed.add(section.id);
			else this.collapsed.delete(section.id);
			this.requestUpdate();
		};
		const body = section.slot
			? html`<slot name=${section.slot}></slot>`
			: html`${repeat(
					section.fields,
					(field) => field.id,
					(field) => fieldTemplate(field, ctx),
				)}${section.note ? html`<p class="note">${this.tr(section.note)}</p>` : nothing}`;
		return html`<section data-chart-format-section=${section.id}
			><button type="button" class="section-toggle" aria-expanded=${String(open)} @click=${toggle}
				>${this.tr(section.title)}</button
			><div class="fields" ?hidden=${!open}>${body}</div></section
		>`;
	}

	override render(): TemplateResult {
		const model = this.model();
		const closeLabel = this.label('close', 'Close');
		const header = html`<header
			><h3>${model ? this.tr(model.title) : ''}</h3
			><button
				type="button"
				class="close"
				data-chart-format-close
				aria-label=${closeLabel}
				title=${closeLabel}
				@click=${() => this.requestClose()}
				>×</button
			></header
		>`;
		if (!model) {
			this.hostLabel = '';
			return html`<div class="panel">${header}</div>`;
		}
		// A tab or mode the target lacks fell back: remember what is showing.
		this.paneTab = model.activeTab;
		this.paneMode = model.mode;
		const title = this.tr(model.title);
		this.hostLabel = title;
		const tags = this.statics().controlTags;
		const picker = this.tr(model.picker);
		const select = unsafeStatic(tags.select);
		const ctx: FieldContext = {
			tags,
			tr: this.tr,
			change: (id, value) => this.edit(id, value),
			preview: this.preview,
			locked: this.disabled,
			autoLabel: this.label('automatic', 'Auto'),
		};
		const textToggle =
			this.sourceValue?.textOptions === false
				? nothing
				: html`<button
						type="button"
						class="text-options"
						data-chart-format-text-options
						aria-pressed=${String(model.mode === 'text')}
						title=${
							model.textOptions
								? this.label('textOptionsHint', 'Fill, outline, effects and text box of the text')
								: this.label('textOptionsNone', 'This element has no text')
						}
						?disabled=${this.disabled || !model.textOptions}
						@click=${() => this.toggleText()}
						>${this.label('textOptions', 'Text Options')}</button
					>`;
		const tabs = model.tabs.length
			? html`<div class="tabs" role="tablist" aria-label=${this.label('tabs', 'Format categories')}
					>${model.tabs.map(
						(tab) =>
							html`<button
								type="button"
								role="tab"
								data-chart-format-tab=${tab.id}
								aria-selected=${String(tab.id === model.activeTab)}
								tabindex=${tab.id === model.activeTab ? 0 : -1}
								@click=${() => {
									this.paneTab = tab.id;
									this.requestUpdate();
								}}
								>${this.tr(tab.label)}</button
							>`,
					)}</div
				>`
			: nothing;
		return html`<div
			class="panel"
			data-chart-format-pane
			data-chart-format-mode=${model.mode}
			data-chart-format-target-key=${model.targetKey}
			@keydown=${this.onKey}
			>${header}<div class="options-row"
				><label class="picker"
					><span>${picker}</span>${staticHtml`<${select}
						data-chart-format-target
						aria-label=${picker}
						data-value=${model.targetKey}
						?disabled=${this.disabled}
						@change=${(event: Event) =>
							this.pickTarget((event.currentTarget as HTMLSelectElement).value)}
					>${model.targets.map(
						(option) =>
							html`<option value=${option.key} .selected=${live(option.key === model.targetKey)}
								>${this.tr(option.label)}</option
							>`,
					)}</${select}>`}</label
				>${textToggle}</div
			>${tabs}<div class="content" role="tabpanel"
				>${repeat(
					model.sections,
					(section) => section.id,
					(section) => this.section(section, ctx),
				)}</div
			></div
		>`;
	}

	protected override updated(changed: PropertyValues): void {
		super.updated(changed);
		// Selects and boxes take their value after their options exist; the box being typed in
		// keeps its text unless an edit was just committed (a refused value shows the model's).
		// The host is the landmark, so a host's slotted section lies inside it.
		this.setAttribute('role', 'complementary');
		if (this.hostLabel) this.setAttribute('aria-label', this.hostLabel);
		else this.removeAttribute('aria-label');
		const force = this.forceSync;
		this.forceSync = false;
		for (const node of this.renderRoot.querySelectorAll<HTMLElement & { value: string }>(
			'[data-value]',
		)) {
			const value = node.dataset.value ?? '';
			const typing =
				node.matches(':focus, :focus-within') && node.value !== (this.synced.get(node) ?? value);
			if (node.value !== value && (force || !typing)) node.value = value;
			this.synced.set(node, value);
		}
	}
}

export const defineChartFormatPane = definer(
	'office-ui-chart-format-pane',
	() => OfficeUiChartFormatPane,
	[defineRadio, defineCheckbox],
);

declare global {
	interface HTMLElementTagNameMap {
		'office-ui-chart-format-pane': OfficeUiChartFormatPane;
	}
}
