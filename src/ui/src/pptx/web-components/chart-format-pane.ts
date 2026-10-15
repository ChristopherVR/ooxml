import type { PptxChartData } from 'ooxml-core/pptx';

import {
	applyChartFormatField,
	buildChartFormatPaneModel,
	translateChartFormatText,
} from '../render';
import type {
	ChartFormatFieldValue,
	ChartFormatPaneModel,
	ChartFormatTabId,
	ChartFormatText,
} from '../render';
import { createPaneField } from './chart-format-pane-fields';
import type { PaneFieldView } from './chart-format-pane-fields';
import { CHART_FORMAT_PANE_STYLES } from './chart-format-pane-styles';
import { attachControlStyles } from './control-styles';
import { attachThemeEditorDock } from './theme-editor-dock';

export type ChartFormatPaneTranslate = (
	key: string,
	params?: Readonly<Record<string, string | number>>,
) => string;

export interface PptxUiChartFormatPaneElement extends HTMLElement {
	/** The selected chart's model. */
	chartData: PptxChartData | undefined;
	/** The element the pane formats (`chartArea`, `series:0`, `axis:1`, ...). */
	target: string;
	translateLabel: ChartFormatPaneTranslate;
	disabled: boolean;
}
/** A committed edit: the whole new chart model, for the host's update + history path. */
export type ChartFormatChangeEvent = CustomEvent<{ chartData: PptxChartData; target: string }>;
/** The element picker chose another element. */
export type ChartFormatTargetEvent = CustomEvent<{ target: string }>;
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-chart-format-pane': PptxUiChartFormatPaneElement;
	}
}

function signature(model: ChartFormatPaneModel): string {
	return JSON.stringify([
		model.title,
		model.targets.map((t) => [t.key, t.label]),
		model.tabs,
		model.activeTab,
		model.sections.map((s) => [
			s.id,
			s.title,
			s.note,
			s.fields.map((f) => [f.id, f.kind, f.label, 'choices' in f ? f.choices : 0]),
		]),
	]);
}

/**
 * `pptx-ui-chart-format-pane`: PowerPoint's chart Format task pane (Format
 * Chart Area, Format Data Series, Format Axis, ...). It owns the view and
 * turns each committed field into a whole new chart model through the shared
 * `applyChartFormatField`; the host applies it and keeps history. Docked at
 * the editor's right edge, a bottom sheet on narrow screens; `inline` lays it
 * out in flow.
 */
export function definePptxChartFormatPane(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-chart-format-pane')) {
		return;
	}
	class ChartFormatPaneElement extends HTMLElement implements PptxUiChartFormatPaneElement {
		#data: PptxChartData | undefined;
		#target = 'chartArea';
		#tab: ChartFormatTabId = 'fillLine';
		#t: ChartFormatPaneTranslate = (key) => key;
		#locked = false;
		#signature = '';
		#fields = new Map<string, PaneFieldView>();
		#collapsed = new Set<string>();
		#undock?: () => void;
		readonly #heading: HTMLHeadingElement;
		readonly #close: HTMLButtonElement;
		readonly #pickerLabel: HTMLSpanElement;
		readonly #picker: HTMLElement & { value: string; disabled: boolean };
		readonly #textOptions: HTMLButtonElement;
		readonly #tabs: HTMLDivElement;
		readonly #content: HTMLDivElement;
		readonly #panel: HTMLDivElement;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachControlStyles(root, CHART_FORMAT_PANE_STYLES);
			const doc = this.ownerDocument;
			this.#panel = doc.createElement('div');
			this.#panel.className = 'panel';
			this.#panel.setAttribute('role', 'complementary');
			this.#panel.dataset.chartFormatPane = '';
			const header = doc.createElement('header');
			this.#heading = doc.createElement('h3');
			this.#close = doc.createElement('button');
			this.#close.type = 'button';
			this.#close.className = 'close';
			this.#close.dataset.chartFormatClose = '';
			this.#close.textContent = '×';
			this.#close.addEventListener('click', () => this.#requestClose());
			header.append(this.#heading, this.#close);
			const options = doc.createElement('div');
			options.className = 'options-row';
			const pickerField = doc.createElement('label');
			pickerField.className = 'picker';
			this.#pickerLabel = doc.createElement('span');
			this.#picker = doc.createElement('pptx-ui-select') as HTMLElement & {
				value: string;
				disabled: boolean;
			};
			this.#picker.dataset.chartFormatTarget = '';
			this.#picker.addEventListener('change', () => this.#pickTarget(this.#picker.value));
			pickerField.append(this.#pickerLabel, this.#picker);
			this.#textOptions = doc.createElement('button');
			this.#textOptions.type = 'button';
			this.#textOptions.className = 'text-options';
			this.#textOptions.disabled = true;
			options.append(pickerField, this.#textOptions);
			this.#tabs = doc.createElement('div');
			this.#tabs.className = 'tabs';
			this.#tabs.setAttribute('role', 'tablist');
			this.#content = doc.createElement('div');
			this.#content.className = 'content';
			this.#content.setAttribute('role', 'tabpanel');
			this.#panel.append(header, options, this.#tabs, this.#content);
			root.append(this.#panel);
			root.addEventListener('keydown', (event) => {
				const key = event as KeyboardEvent;
				key.stopPropagation();
				if (key.key === 'Escape' && !key.defaultPrevented) {
					key.preventDefault();
					this.#requestClose();
				}
			});
		}
		get chartData(): PptxChartData | undefined {
			return this.#data;
		}
		set chartData(value: PptxChartData | undefined) {
			this.#data = value;
			this.#paint();
		}
		get target(): string {
			return this.#target;
		}
		set target(value: string) {
			this.#target = value || 'chartArea';
			this.#paint();
		}
		get translateLabel(): ChartFormatPaneTranslate {
			return this.#t;
		}
		set translateLabel(value: ChartFormatPaneTranslate) {
			this.#t = value;
			this.#signature = '';
			this.#paint();
		}
		get disabled(): boolean {
			return this.#locked;
		}
		set disabled(value: boolean) {
			this.#locked = value;
			this.#paint();
		}
		connectedCallback(): void {
			this.#undock = attachThemeEditorDock(this);
			this.#paint();
		}
		disconnectedCallback(): void {
			this.#undock?.();
		}
		#tr = (text: ChartFormatText): string => translateChartFormatText(text, this.#t);
		#label(key: string, fallback: string): string {
			return this.#tr({ key, fallback });
		}
		#requestClose(): void {
			this.dispatchEvent(new CustomEvent('chart-format-close', { bubbles: true, composed: true }));
		}
		#pickTarget(target: string): void {
			this.#target = target;
			this.#paint();
			this.dispatchEvent(
				new CustomEvent('chart-format-target', {
					detail: { target },
					bubbles: true,
					composed: true,
				}),
			);
		}
		#change(id: string, value: ChartFormatFieldValue): void {
			if (this.#locked || !this.#data) {
				return;
			}
			const next = applyChartFormatField(this.#data, this.#target, id, value);
			if (!next) {
				this.#signature = '';
				this.#paint();
				return;
			}
			this.#data = next;
			this.#paint();
			this.dispatchEvent(
				new CustomEvent('chart-format-change', {
					detail: { chartData: next, target: this.#target },
					bubbles: true,
					composed: true,
				}),
			);
		}
		#paint(): void {
			if (!this.#data) {
				this.#heading.textContent = '';
				this.#content.replaceChildren();
				this.#signature = '';
				return;
			}
			const model = buildChartFormatPaneModel(this.#data, this.#target, this.#tab);
			this.#tab = model.activeTab;
			const title = this.#tr(model.title);
			this.#heading.textContent = title;
			this.#panel.setAttribute('aria-label', title);
			this.#panel.dataset.chartFormatTargetKey = model.targetKey;
			this.#close.setAttribute('aria-label', this.#label('pptx.chartFormat.close', 'Close'));
			this.#close.title = this.#close.getAttribute('aria-label') ?? '';
			this.#tabs.setAttribute(
				'aria-label',
				this.#label('pptx.chartFormat.tabs', 'Format categories'),
			);
			const next = signature(model);
			if (next !== this.#signature) {
				this.#signature = next;
				this.#build(model);
			}
			this.#picker.value = model.targetKey;
			this.#picker.disabled = this.#locked;
			for (const section of model.sections) {
				for (const field of section.fields) {
					this.#fields.get(field.id)?.paint(field, this.#locked);
				}
			}
		}
		#build(model: ChartFormatPaneModel): void {
			const doc = this.ownerDocument;
			this.#pickerLabel.textContent = this.#label('pptx.chartFormat.chartOptions', 'Chart Options');
			this.#picker.setAttribute('aria-label', this.#pickerLabel.textContent);
			this.#textOptions.textContent = this.#label('pptx.chartFormat.textOptions', 'Text Options');
			this.#textOptions.title = this.#label(
				'pptx.chartFormat.textOptionsHint',
				'Format chart text with Home > Font',
			);
			this.#picker.replaceChildren(
				...model.targets.map((option) => {
					const node = doc.createElement('option');
					node.value = option.key;
					node.textContent = this.#tr(option.label);
					return node;
				}),
			);
			this.#tabs.replaceChildren(
				...model.tabs.map((tab) => {
					const button = doc.createElement('button');
					button.type = 'button';
					button.setAttribute('role', 'tab');
					button.dataset.chartFormatTab = tab.id;
					button.setAttribute('aria-selected', String(tab.id === model.activeTab));
					button.tabIndex = tab.id === model.activeTab ? 0 : -1;
					button.textContent = this.#tr(tab.label);
					button.addEventListener('click', () => {
						this.#tab = tab.id;
						this.#paint();
					});
					return button;
				}),
			);
			this.#fields.clear();
			const autoLabel = this.#label('pptx.chartFormat.automatic', 'Auto');
			this.#content.replaceChildren(
				...model.sections.map((section) => {
					const node = doc.createElement('section');
					node.dataset.chartFormatSection = section.id;
					const toggle = doc.createElement('button');
					toggle.type = 'button';
					toggle.className = 'section-toggle';
					const open = !this.#collapsed.has(section.id);
					toggle.setAttribute('aria-expanded', String(open));
					toggle.textContent = this.#tr(section.title);
					const body = doc.createElement('div');
					body.className = 'fields';
					body.hidden = !open;
					toggle.addEventListener('click', () => {
						const expanded = body.hidden;
						body.hidden = !expanded;
						toggle.setAttribute('aria-expanded', String(expanded));
						if (expanded) this.#collapsed.delete(section.id);
						else this.#collapsed.add(section.id);
					});
					for (const field of section.fields) {
						const view = createPaneField(
							doc,
							field,
							this.#tr,
							(id, v) => this.#change(id, v),
							autoLabel,
						);
						this.#fields.set(field.id, view);
						body.append(view.node);
					}
					if (section.note) {
						const note = doc.createElement('p');
						note.className = 'note';
						note.textContent = this.#tr(section.note);
						body.append(note);
					}
					node.append(toggle, body);
					return node;
				}),
			);
		}
	}
	registry.define('pptx-ui-chart-format-pane', ChartFormatPaneElement);
}
