// SmartArt in the Word editor: Insert > SmartArt, the SmartArt Design commands and the shared text
// pane (`SmartArtTextPaneController`, which Excel uses too), wired to the core's editing
// operations. The layout, colour and quick-style definitions a new or re-laid-out diagram needs
// come from `ooxml-core/diagram/builtin`, loaded on first use (its 176 layouts are large), and are
// staged as pending media for the save.
import type { EditorView } from 'prosemirror-view';
import type { DiagramInsertPreset } from 'ooxml-core/diagram';
import {
	changeDocxDiagramLayout,
	createDocxDiagram,
	formatDocxDiagram,
	insertedDiagramPartNames,
	type PendingMediaPart,
} from 'ooxml-core/docx';
import { SmartArtTextPaneController } from '../smartart';
import type { RibbonAction } from './ribbon-action';
import { focusView } from './focus-view';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import {
	insertDiagram,
	replaceDiagram,
	runDiagramEdit,
	selectedDiagram,
	setDiagramNodeText,
} from './smartart-commands';
import { createSmartArtDialog, type SmartArtDialog } from './smartart-dialog';

type SmartArtKey = Extract<RibbonAction, { type: 'smartArt' }>['key'];

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
	private readonly text: SmartArtTextPaneController;
	private locale: EditorLocale = 'en';
	/** Whether the last pane edit came from a ribbon command (focus returns to the document). */
	private refocus = false;

	constructor(private readonly host: SmartArtHost) {
		this.dialog = createSmartArtDialog(() => focusView(host.view()));
		this.text = new SmartArtTextPaneController(
			document,
			{
				nodes: () => {
					const view = host.view();
					return view ? selectedDiagram(view.state)?.diagram.nodes : undefined;
				},
				edit: (edit) => {
					const view = host.view();
					const nodeId = view ? runDiagramEdit(view, edit) : undefined;
					if (this.refocus) focusView(view);
					return nodeId;
				},
				setText: (id, text) => {
					const view = host.view();
					if (view) setDiagramNodeText(view, id, text);
				},
				toggled: (open) => {
					host.refresh();
					if (!open) focusView(host.view());
				},
			},
			'dve-smartart-text-pane',
		);
	}

	get pane(): HTMLElement {
		return this.text.pane;
	}

	get textPaneOpen(): boolean {
		return this.text.open;
	}

	/** Handles the SmartArt actions; returns false for everything else. */
	handle(action: RibbonAction): boolean {
		if (action.type === 'smartArtFormat') {
			const view = this.host.view();
			const selected = view && selectedDiagram(view.state);
			if (view?.editable && selected) {
				const next = formatDocxDiagram(selected.diagram, action.change);
				if (next !== selected.diagram) replaceDiagram(view, selected.pos, next);
			}
			return true;
		}
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
		if (!view || !selectedDiagram(view.state)) return;
		if (key === 'textPane') {
			this.text.setOpen(!this.text.open);
			return;
		}
		if (key === 'changeLayout') {
			this.dialog.open((preset) => void this.relayout(preset), undefined);
			return;
		}
		this.refocus = true;
		try {
			this.text.runCommand(key);
		} finally {
			this.refocus = false;
		}
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
				layoutSource: {
					layoutXml: builtin.loadBuiltinSmartArtLayoutXml(layoutId),
					colorsXml: builtin.BUILTIN_SMARTART_COLORS_XML,
					quickStyleXml: builtin.BUILTIN_SMARTART_QUICK_STYLE_XML,
				},
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
			this.text.select(diagram.nodes[0]?.id ?? null);
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
			const next = changeDocxDiagramLayout(
				selected.diagram,
				layoutId,
				preset.family,
				builtin.loadBuiltinSmartArtLayoutXml(layoutId),
			);
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
		this.text.sync();
	}

	setLocale(locale: EditorLocale): void {
		this.locale = locale;
		this.dialog.setLocale(locale);
		this.text.setLabels((key) => translate(locale, key as LocalizationKey));
	}

	/** A new document: no current node, and the text pane closed. */
	reset(): void {
		if (this.dialog.isOpen) this.dialog.close();
		this.text.reset();
	}
}
