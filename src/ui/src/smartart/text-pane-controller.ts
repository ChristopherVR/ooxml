// The wiring every product's SmartArt text pane shares: it creates `<office-ui-smartart-text-pane>`,
// turns its requests (typing, Enter, Tab, Shift+Tab, Backspace) into the shared node commands of
// `ooxml-core/diagram` and runs the SmartArt Design node commands (`diagramDesignEdit`) on the
// node the pane is on. The product supplies the selected diagram's nodes and applies edits through
// its own undo history (Word: a ProseMirror transaction; Excel: an edit session step).
import {
	diagramDesignEdit,
	diagramNodeOutline,
	type DiagramDesignCommandKey,
	type DiagramHostNode,
	type DiagramNodeEdit,
} from 'ooxml-core/diagram';
import {
	defineSmartArtTextPane,
	type OfficeSmartArtNodeRequest,
	type OfficeSmartArtTextChange,
	type OfficeSmartArtTextNode,
	type OfficeSmartArtTextPaneLabels,
} from './smartart-text-pane';

export interface SmartArtTextPaneHost {
	/** The selected diagram's nodes in tree order, or undefined when no SmartArt is selected. */
	nodes(): readonly DiagramHostNode[] | undefined;
	/**
	 * Applies one command to the selected diagram. Returns the node it created or acted on, `''`
	 * when it changed the diagram without one, and undefined when nothing changed.
	 */
	edit(edit: DiagramNodeEdit): string | undefined;
	/** Typing in a line (hosts may join keystrokes into one undo step). */
	setText(id: string, text: string): void;
	/** The pane opened or closed (the host refreshes its ribbon and focus). */
	toggled(open: boolean): void;
}

type TextPane = HTMLElement & {
	nodes: OfficeSmartArtTextNode[];
	selected: string | null;
	closable: boolean;
	labels: Partial<OfficeSmartArtTextPaneLabels>;
	focusSelected(): void;
};

/** One product's text pane and the node the SmartArt Design commands act on. */
export class SmartArtTextPaneController {
	readonly pane: TextPane;
	private paneOpen = false;
	/** The node the text pane (and so Add Shape, Promote...) acts on. */
	private current: string | null = null;

	constructor(
		doc: Document,
		private readonly host: SmartArtTextPaneHost,
		className: string,
	) {
		defineSmartArtTextPane(doc.defaultView?.customElements);
		this.pane = doc.createElement('office-ui-smartart-text-pane') as TextPane;
		this.pane.className = className;
		this.pane.closable = true;
		this.pane.hidden = true;
		this.pane.addEventListener('office-smartart-text-change', (event) => {
			const { id, text } = (event as CustomEvent<OfficeSmartArtTextChange>).detail;
			host.setText(id, text);
		});
		const request = (name: string, edit: (id: string) => DiagramNodeEdit) =>
			this.pane.addEventListener(name, (event) => {
				const { id } = (event as CustomEvent<OfficeSmartArtNodeRequest>).detail;
				this.current = id;
				this.apply(edit(id));
			});
		request('office-smartart-node-add', (id) => ({ kind: 'add', position: 'after', nodeId: id }));
		request('office-smartart-node-promote', (id) => ({ kind: 'promote', nodeId: id }));
		request('office-smartart-node-demote', (id) => ({ kind: 'demote', nodeId: id }));
		this.pane.addEventListener('office-smartart-node-remove', (event) => {
			const { id } = (event as CustomEvent<OfficeSmartArtNodeRequest>).detail;
			const nodes = this.pane.nodes;
			const index = nodes.findIndex((node) => node.id === id);
			this.current = nodes[index - 1]?.id ?? nodes[index + 1]?.id ?? null;
			this.apply({ kind: 'remove', nodeId: id });
		});
		this.pane.addEventListener('office-smartart-node-select', (event) => {
			this.current = (event as CustomEvent<OfficeSmartArtNodeRequest>).detail.id;
		});
		this.pane.addEventListener('office-smartart-text-pane-close', () => this.setOpen(false));
	}

	get open(): boolean {
		return this.paneOpen;
	}

	/** Opens or closes the pane (Text Pane toggle); focuses its current line on opening. */
	setOpen(open: boolean): void {
		if (this.paneOpen === open) return;
		this.paneOpen = open;
		this.sync();
		this.host.toggled(open);
		if (open) this.pane.focusSelected();
	}

	/** Puts the pane on a node (a newly inserted diagram's first node, for example). */
	select(nodeId: string | null): void {
		this.current = nodeId;
		this.sync();
	}

	/** The node commands act on: the text pane's, else the diagram's last node. */
	currentNode(): string | undefined {
		const nodes = this.host.nodes();
		if (!nodes) return undefined;
		if (this.current && nodes.some((node) => node.id === this.current)) return this.current;
		return nodes.at(-1)?.id;
	}

	/**
	 * Runs a SmartArt Design node command on the current node; returns false for the commands that
	 * are not node edits (`textPane`, `changeLayout`) or when nothing is selected.
	 */
	runCommand(key: DiagramDesignCommandKey): boolean {
		if (!this.host.nodes()) return false;
		const edit = diagramDesignEdit(key, this.currentNode());
		if (!edit) return false;
		this.apply(edit);
		return true;
	}

	private apply(edit: DiagramNodeEdit): void {
		const nodeId = this.host.edit(edit);
		if (nodeId) this.current = nodeId;
		this.sync();
	}

	/** Shows the pane for the selected diagram while it is open, with its current outline. */
	sync(): void {
		const nodes = this.host.nodes();
		this.pane.hidden = !this.paneOpen || !nodes;
		if (!nodes) return;
		this.pane.nodes = diagramNodeOutline(nodes);
		const current = this.currentNode() ?? null;
		if (this.pane.selected !== current) this.pane.selected = current;
	}

	/** Translated pane strings; `t` looks up the shared SmartArt UI keys. */
	setLabels(t: (key: string) => string): void {
		this.pane.labels = {
			title: t('Type your text here'),
			placeholder: t('[Text]'),
			close: t('Close'),
			outline: t('SmartArt text'),
			empty: t('This diagram has no shapes with text.'),
		};
	}

	/** A new document: no current node, and the pane closed. */
	reset(): void {
		this.current = null;
		this.paneOpen = false;
		this.sync();
	}
}
