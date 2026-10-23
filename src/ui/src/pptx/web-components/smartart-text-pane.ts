import type { PptxSmartArtData, PptxSmartArtNodeStyle } from 'ooxml-core/pptx';
import { defineSmartArtTextPane } from '../../smartart/smartart-text-pane';
import type {
	OfficeSmartArtNodeMove,
	OfficeSmartArtNodeRequest,
	OfficeSmartArtNodeStyleChange,
	OfficeSmartArtTextChange,
	OfficeSmartArtTextPaneLabels,
	OfficeUiSmartArtTextPane,
} from '../../smartart/smartart-text-pane';
import { applySmartArtTextPaneRequest, smartArtTextPaneView } from '../render';
import type { SmartArtTextPaneRequest } from '../render';
import { base, bridged } from './office-aliases';

type Translate = (key: string, params?: Record<string, string | number>) => string;

/** What a binding passes to `pptx-ui-smartart-text-pane` through `state`. */
export interface SmartArtTextPaneViewState {
	smartArtData: PptxSmartArtData | undefined;
	canEdit: boolean;
	translate?: Translate;
}

/**
 * Detail of `pptx-smartart-data-change`: the SmartArt data after the edit, ready to commit, and
 * the colour a style edit picked (for the recent colours list).
 */
export interface SmartArtTextPaneChange {
	smartArtData: PptxSmartArtData;
	recentColor?: string;
}
export type SmartArtTextPaneChangeEvent = CustomEvent<SmartArtTextPaneChange>;

export interface PptxUiSmartArtTextPaneElement extends OfficeUiSmartArtTextPane {
	state: SmartArtTextPaneViewState;
}
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-smartart-text-pane': PptxUiSmartArtTextPaneElement;
	}
}

/** The pane's text from the pptx dictionary (the keys every binding's text pane used). */
export function smartArtTextPaneLabels(t: Translate, count: number): OfficeSmartArtTextPaneLabels {
	return {
		title: `${t('pptx.smartart.textPane')} (${count})`,
		placeholder: t('pptx.smartArt.nodePlaceholder'),
		close: t('pptx.smartart.close'),
		outline: t('pptx.smartart.textPane'),
		empty: t('pptx.smartart.textPaneEmpty'),
		add: t('pptx.smartart.addItem'),
		moveUp: t('pptx.smartArt.moveUp'),
		moveDown: t('pptx.smartArt.moveDown'),
		addChild: t('pptx.smartArt.addSubItem'),
		remove: t('pptx.smartArt.remove'),
		removeUnavailable: t('pptx.smartArt.layoutMinimum'),
		fill: t('pptx.smartArt.fill'),
		fontColor: t('pptx.smartArt.text'),
		bold: t('pptx.textPanel.bold'),
		italic: t('pptx.textPanel.italic'),
	};
}

const identity: Translate = (key) => key;

/** The element events and the pane request each one makes. */
const REQUESTS: Record<string, (detail: never) => SmartArtTextPaneRequest> = {
	'office-smartart-text-change': (d: OfficeSmartArtTextChange) => ({ kind: 'text', ...d }),
	'office-smartart-node-add': (d: OfficeSmartArtNodeRequest) => ({ kind: 'add', id: d.id }),
	'office-smartart-node-remove': (d: OfficeSmartArtNodeRequest) => ({ kind: 'remove', id: d.id }),
	'office-smartart-node-promote': (d: OfficeSmartArtNodeRequest) => ({ kind: 'promote', id: d.id }),
	'office-smartart-node-demote': (d: OfficeSmartArtNodeRequest) => ({ kind: 'demote', id: d.id }),
	'office-smartart-node-add-child': (d: OfficeSmartArtNodeRequest) => ({
		kind: 'addChild',
		id: d.id,
	}),
	'office-smartart-node-move': (d: OfficeSmartArtNodeMove) => ({ kind: 'move', ...d }),
	'office-smartart-node-append': () => ({ kind: 'append' }),
	'office-smartart-node-style': (d: OfficeSmartArtNodeStyleChange) => ({
		kind: 'style',
		id: d.id,
		style: d.style as Partial<PptxSmartArtNodeStyle>,
	}),
};

/**
 * `pptx-ui-smartart-text-pane`: the shared SmartArt text pane fed a pptx diagram, with
 * PowerPoint's inspector tools (append, move, sub-item, remove, per-node style). A binding sets
 * `state` and commits the `pptx-smartart-data-change` detail's `smartArtData` through its own
 * update path; the edits go through the pptx SmartArt facades (`smartart-text-pane-model`), so
 * every binding edits a diagram the same way.
 */
export function definePptxSmartArtTextPane(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-smartart-text-pane')) {
		return;
	}
	const Base = bridged(
		base(registry, defineSmartArtTextPane, 'office-ui-smartart-text-pane'),
		'pptx-ui-smartart-text-pane',
	) as unknown as new () => OfficeUiSmartArtTextPane;
	class PptxSmartArtTextPane extends Base {
		#model: SmartArtTextPaneViewState = { smartArtData: undefined, canEdit: false };
		constructor() {
			super();
			this.tools = true;
			for (const [name, toRequest] of Object.entries(REQUESTS)) {
				this.addEventListener(name, (event) =>
					this.#apply(toRequest((event as CustomEvent).detail as never)),
				);
			}
		}
		get state(): SmartArtTextPaneViewState {
			return this.#model;
		}
		set state(value: SmartArtTextPaneViewState) {
			this.#model = value;
			this.disabled = !value.canEdit;
			this.#render();
		}
		#render(): void {
			const data = this.#model.smartArtData;
			const t = this.#model.translate ?? identity;
			this.hidden = !data;
			const view = data
				? smartArtTextPaneView(data, t)
				: { lines: [], appendDisabled: true, notes: [] };
			this.labels = smartArtTextPaneLabels(t, view.lines.length);
			this.appendDisabled = view.appendDisabled;
			this.notes = view.notes;
			this.nodes = view.lines;
		}
		#apply(request: SmartArtTextPaneRequest): void {
			const current = this.#model.smartArtData;
			if (!current || !this.#model.canEdit) {
				return;
			}
			const t = this.#model.translate ?? identity;
			const edit = applySmartArtTextPaneRequest(current, request, t('pptx.smartart.subItemShort'));
			if (!edit) {
				return;
			}
			// Later requests build on this edit even before the host sets `state` again.
			this.#model = { ...this.#model, smartArtData: edit.data };
			if (edit.focusNodeId) {
				this.selected = edit.focusNodeId;
			}
			this.#render();
			const color =
				request.kind === 'style' ? (request.style.fillColor ?? request.style.fontColor) : undefined;
			const detail: SmartArtTextPaneChange = {
				smartArtData: edit.data,
				...(color ? { recentColor: color } : {}),
			};
			this.fire('pptx-smartart-data-change', detail);
		}
	}
	registry.define(
		'pptx-ui-smartart-text-pane',
		PptxSmartArtTextPane as unknown as CustomElementConstructor,
	);
}
