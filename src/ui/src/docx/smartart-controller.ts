// SmartArt in the Word editor: Insert > SmartArt, the SmartArt Design commands and the shared text
// pane, wired to the core's editing operations. The layout, colour and quick-style definitions a
// new or re-laid-out diagram needs come from `ooxml-core/diagram/builtin`, loaded on first use
// (its 176 layouts are large), and are staged as pending media for the save.
import type { EditorView } from 'prosemirror-view';
import type { DiagramInsertPreset } from 'ooxml-core/diagram';
import {
	changeDocxDiagramLayout,
	createDocxDiagram,
	insertedDiagramPartNames,
	type DocxDiagramEdit,
	type PendingMediaPart,
} from 'ooxml-core/docx';
import {
	defineSmartArtTextPane,
	type OfficeSmartArtNodeRequest,
	type OfficeSmartArtTextChange,
	type OfficeSmartArtTextNode,
	type OfficeSmartArtTextPaneLabels,
} from '../smartart';
import type { RibbonAction } from './ribbon-action';
import { focusView } from './focus-view';
import { translate, type EditorLocale } from './localization';
import {
	diagramOutline,
	insertDiagram,
	replaceDiagram,
	runDiagramEdit,
	selectedDiagram,
	setDiagramNodeText,
} from './smartart-commands';
import { createSmartArtDialog, type SmartArtDialog } from './smartart-dialog';

type SmartArtKey = Extract<RibbonAction, { type: 'smartArt' }>['key'];
type TextPane = HTMLElement & {
	nodes: OfficeSmartArtTextNode[];
	selected: string | null;
	closable: boolean;
	labels: Partial<OfficeSmartArtTextPaneLabels>;
	focusSelected(): void;
};

export interface SmartArtHost {
	view(): EditorView | undefined;
	/** The editable content width in CSS pixels; a new diagram is fitted to it. */
	contentWidth(): number;
	stage(partName: string, part: PendingMediaPart): void;
	reportError(error: Error): void;
	refresh(): void;
}

/** Word's default SmartArt size, 6 by 3.5 inches. */
const DEFAULT_SIZE = { width: 576, height: 336 };
let serial = 0;

const xmlPart = (xml: string): PendingMediaPart => ({
	bytes: new TextEncoder().encode(xml),
	contentType: 'application/xml',
});

export class SmartArtController {
	readonly dialog: SmartArtDialog;
	readonly pane: TextPane;
	private paneOpen = false;
	/** The node the text pane (and so Add Shape, Promote...) acts on, per diagram. */
	private current: string | null = null;
	private locale: EditorLocale = 'en';

	constructor(private readonly host: SmartArtHost) {
		this.dialog = createSmartArtDialog(() => focusView(host.view()));
		defineSmartArtTextPane();
		this.pane = document.createElement('office-ui-smartart-text-pane') as TextPane;
		this.pane.className = 'dve-smartart-text-pane';
		this.pane.closable = true;
		this.pane.hidden = true;
		this.pane.addEventListener('office-smartart-text-change', (event) => {
			const { id, text } = (event as CustomEvent<OfficeSmartArtTextChange>).detail;
			const view = host.view();
			if (view) setDiagramNodeText(view, id, text);
		});
		const request = (name: string, edit: (id: string) => DocxDiagramEdit) =>
			this.pane.addEventListener(name, (event) => {
				const { id } = (event as CustomEvent<OfficeSmartArtNodeRequest>).detail;
				this.current = id;
				this.edit(edit(id), false);
			});
		request('office-smartart-node-add', (id) => ({ kind: 'add', position: 'after', nodeId: id }));
		request('office-smartart-node-promote', (id) => ({ kind: 'promote', nodeId: id }));
		request('office-smartart-node-demote', (id) => ({ kind: 'demote', nodeId: id }));
		this.pane.addEventListener('office-smartart-node-remove', (event) => {
			const { id } = (event as CustomEvent<OfficeSmartArtNodeRequest>).detail;
			const nodes = this.pane.nodes;
			const index = nodes.findIndex((node) => node.id === id);
			this.current = nodes[index - 1]?.id ?? nodes[index + 1]?.id ?? null;
			this.edit({ kind: 'remove', nodeId: id }, false);
		});
		this.pane.addEventListener('office-smartart-node-select', (event) => {
			this.current = (event as CustomEvent<OfficeSmartArtNodeRequest>).detail.id;
		});
		this.pane.addEventListener('office-smartart-text-pane-close', () => {
			this.paneOpen = false;
			this.sync();
			host.refresh();
			focusView(host.view());
		});
	}

	get textPaneOpen(): boolean {
		return this.paneOpen;
	}

	/** Handles the SmartArt actions; returns false for everything else. */
	handle(action: RibbonAction): boolean {
		if (action.type === 'insertSmartArt') {
			if (this.host.view()?.editable) this.dialog.open((preset) => void this.insert(preset));
			return true;
		}
		if (action.type !== 'smartArt') return false;
		this.run(action.key);
		return true;
	}

	private run(key: SmartArtKey): void {
		const view = this.host.view();
		const selected = view && selectedDiagram(view.state);
		if (!view || !selected) return;
		const nodeId = this.currentNode();
		if (key === 'textPane') {
			this.paneOpen = !this.paneOpen;
			this.sync();
			this.host.refresh();
			if (this.paneOpen) this.pane.focusSelected();
			return;
		}
		if (key === 'changeLayout') {
			this.dialog.open((preset) => void this.relayout(preset), undefined);
			return;
		}
		const edits: Record<Exclude<SmartArtKey, 'textPane' | 'changeLayout'>, DocxDiagramEdit> = {
			addAfter: { kind: 'add', position: 'after', ...(nodeId ? { nodeId } : {}) },
			addBefore: { kind: 'add', position: 'before', ...(nodeId ? { nodeId } : {}) },
			addAbove: { kind: 'add', position: 'above', ...(nodeId ? { nodeId } : {}) },
			addBelow: { kind: 'add', position: 'below', ...(nodeId ? { nodeId } : {}) },
			promote: { kind: 'promote', nodeId: nodeId ?? '' },
			demote: { kind: 'demote', nodeId: nodeId ?? '' },
			moveUp: { kind: 'move', nodeId: nodeId ?? '', direction: -1 },
			moveDown: { kind: 'move', nodeId: nodeId ?? '', direction: 1 },
		};
		this.edit(edits[key], true);
	}

	/** The node commands act on: the text pane's, else the diagram's last node. */
	private currentNode(): string | undefined {
		const view = this.host.view();
		const selected = view && selectedDiagram(view.state);
		if (!selected) return undefined;
		const nodes = selected.diagram.nodes;
		if (this.current && nodes.some((node) => node.id === this.current)) return this.current;
		return nodes.at(-1)?.id;
	}

	private edit(edit: DocxDiagramEdit, refocus: boolean): void {
		const view = this.host.view();
		if (!view) return;
		const nodeId = runDiagramEdit(view, edit);
		if (nodeId) this.current = nodeId;
		this.sync();
		if (refocus) focusView(view);
	}

	private async definitions() {
		return import('ooxml-core/diagram/builtin');
	}

	private async insert(preset: DiagramInsertPreset): Promise<void> {
		const view = this.host.view();
		if (!view?.editable) return;
		try {
			const builtin = await this.definitions();
			const layoutId = builtin.builtinLayoutIdForPreset(preset.layout);
			if (!layoutId)
				throw new Error(translate(this.locale, 'The SmartArt layouts could not be loaded.'));
			const width = Math.min(DEFAULT_SIZE.width, Math.max(96, this.host.contentWidth()));
			const height = Math.round((DEFAULT_SIZE.height * width) / DEFAULT_SIZE.width);
			const partSuffix = `-dve-${Date.now().toString(36)}-${++serial}`;
			const diagram = createDocxDiagram({
				layoutId,
				family: preset.family,
				colorsId: builtin.BUILTIN_SMARTART_COLORS_ID,
				quickStyleId: builtin.BUILTIN_SMARTART_QUICK_STYLE_ID,
				// Word starts every shape empty, showing its "[Text]" placeholder.
				items: preset.defaultItems.map(() => ''),
				widthPx: width,
				heightPx: height,
				partSuffix,
			});
			const names = insertedDiagramPartNames(partSuffix);
			this.host.stage(names.layout, xmlPart(builtin.loadBuiltinSmartArtLayoutXml(layoutId)));
			this.host.stage(names.colors, xmlPart(builtin.BUILTIN_SMARTART_COLORS_XML));
			this.host.stage(names.quickStyle, xmlPart(builtin.BUILTIN_SMARTART_QUICK_STYLE_XML));
			insertDiagram(view, {
				relId: '',
				partName: '',
				contentType: 'application/octet-stream',
				widthPx: width,
				heightPx: height,
				unsupported: 'SmartArt',
				diagram,
			});
			this.current = diagram.nodes[0]?.id ?? null;
			this.sync();
			focusView(view);
		} catch (error) {
			this.host.reportError(error instanceof Error ? error : new Error(String(error)));
		}
	}

	private async relayout(preset: DiagramInsertPreset): Promise<void> {
		const view = this.host.view();
		const selected = view && selectedDiagram(view.state);
		const layoutPart = selected?.diagram.parts.layout?.partName;
		if (!view?.editable || !selected || !layoutPart) return;
		try {
			const builtin = await this.definitions();
			const layoutId = builtin.builtinLayoutIdForPreset(preset.layout);
			if (!layoutId)
				throw new Error(translate(this.locale, 'The SmartArt layouts could not be loaded.'));
			const next = changeDocxDiagramLayout(selected.diagram, layoutId, preset.family);
			if (next === selected.diagram) return;
			this.host.stage(layoutPart, xmlPart(builtin.loadBuiltinSmartArtLayoutXml(layoutId)));
			replaceDiagram(view, selected.pos, next);
			this.sync();
			focusView(view);
		} catch (error) {
			this.host.reportError(error instanceof Error ? error : new Error(String(error)));
		}
	}

	/** Shows the text pane for the selected diagram while it is open. */
	sync(): void {
		const view = this.host.view();
		const selected = view && selectedDiagram(view.state);
		this.pane.hidden = !this.paneOpen || !selected;
		if (!selected) return;
		const outline = diagramOutline(selected.diagram);
		this.pane.nodes = outline;
		const current = this.currentNode();
		if (this.pane.selected !== (current ?? null)) this.pane.selected = current ?? null;
	}

	setLocale(locale: EditorLocale): void {
		this.locale = locale;
		this.dialog.setLocale(locale);
		const t = (key: Parameters<typeof translate>[1]) => translate(locale, key);
		this.pane.labels = {
			title: t('Type your text here'),
			placeholder: t('[Text]'),
			close: t('Close'),
			outline: t('SmartArt text'),
			empty: t('This diagram has no shapes with text.'),
		};
	}

	/** A new document: no current node, and the text pane closed. */
	reset(): void {
		this.current = null;
		this.paneOpen = false;
		if (this.dialog.isOpen) this.dialog.close();
		this.sync();
	}
}
