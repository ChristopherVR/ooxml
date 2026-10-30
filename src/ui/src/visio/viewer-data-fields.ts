import type { VisioPage, VisioShape } from 'ooxml-core/visio';
import {
	visioDataGraphicItems,
	visioDeleteDataGraphicItemEdits,
	visioSetDataGraphicItemEdits,
} from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController, type ViewerState } from './controller';

type Run = (action: () => Promise<void>, success: string) => void;

/** Fields a person can show: visible Shape Data rows, never this viewer's own marker rows. */
const fieldsOf = (shape: VisioShape) =>
	(shape.shapeData ?? []).filter((row) => !row.invisible && !row.name.startsWith('_'));

/**
 * Visio's Data Graphic Fields pane: the Shape Data fields of the selected shape, each with a
 * check box that shows the value on the shape as a text callout and hides it again.
 */
export class ViewerGraphicFields {
	readonly pane: HTMLElement;
	#list: HTMLElement;
	#hint: HTMLElement;
	#signature = '';
	constructor(
		root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly run: Run,
	) {
		const doc = root.ownerDocument;
		this.pane = doc.createElement('office-ui-task-pane');
		this.pane.className = 'graphic-fields';
		this.pane.hidden = true;
		this.pane.setAttribute('label', 'Data Graphic Fields');
		this.pane.setAttribute('close-label', 'Close Data Graphic Fields');
		this.#hint = doc.createElement('p');
		this.#hint.className = 'graphic-fields-hint';
		this.#list = doc.createElement('div');
		this.#list.className = 'graphic-fields-list';
		this.#list.setAttribute('role', 'group');
		this.#list.setAttribute('aria-label', 'Fields to show as text');
		this.pane.append(this.#hint, this.#list);
		this.pane.addEventListener('office-pane-close', () => this.toggle(false));
		(root.querySelector('.workspace') ?? root).append(this.pane);
		this.#list.addEventListener('change', (event) => {
			const box = event.target as HTMLInputElement;
			if (box.dataset.field !== undefined) this.#toggle(box.dataset.field, box.checked);
		});
	}
	toggle(open = this.pane.hidden): void {
		this.pane.hidden = !open;
		this.#signature = '';
		this.render(this.controller.state);
	}
	#page(state: ViewerState): VisioPage | undefined {
		return state.document?.pages[state.pageIndex];
	}
	#targets(state: ViewerState, page: VisioPage): VisioShape[] {
		const ids = new Set(state.selectedShapes.map((shape) => shape.id));
		return page.shapes.filter((shape) => ids.has(shape.id));
	}
	#toggle(field: string, on: boolean): void {
		const state = this.controller.state;
		const page = this.#page(state);
		if (!page || !canEditDiagram(state)) return;
		const owners = this.#targets(state, page).filter((shape) =>
			fieldsOf(shape).some((row) => row.name === field),
		);
		const ids = owners.map((shape) => shape.id);
		const item = { kind: 'text', field } as const;
		const edits = on
			? visioSetDataGraphicItemEdits(page, ids, item)
			: visioDeleteDataGraphicItemEdits(page, ids, item);
		if (edits.length)
			this.run(
				() => this.controller.applyEdits(edits),
				on ? `Showing ${field} on the selection.` : `Hid ${field} on the selection.`,
			);
	}
	render(state: ViewerState): void {
		const box = this.pane.getRootNode() as ShadowRoot;
		const check = box.querySelector<HTMLElement & { checked: boolean; disabled: boolean }>(
			'[data-check="data-graphic-fields"]',
		);
		if (check) {
			check.checked = !this.pane.hidden;
			check.disabled = !state.document;
		}
		if (this.pane.hidden) return;
		const page = this.#page(state);
		const editable = canEditDiagram(state) && !state.loading && !state.edit.busy;
		const owners = page ? this.#targets(state, page) : [];
		const first = owners[0];
		const rows = first ? fieldsOf(first) : [];
		const on = new Set(
			first && page
				? visioDataGraphicItems(page, first.id)
						.filter((item) => item.kind === 'text')
						.map((item) => item.field)
				: [],
		);
		const signature = JSON.stringify([
			first?.id,
			editable,
			rows.map((row) => [row.name, row.label, row.value, on.has(row.name)]),
		]);
		if (signature === this.#signature) return;
		this.#signature = signature;
		this.#hint.textContent = !first
			? 'Select a shape to choose which of its data fields show as text.'
			: rows.length
				? `Fields of ${first.text.plainText || `shape ${first.id}`}. Checked fields show as text on the shape.`
				: 'The selected shape has no Shape Data fields. Link data or define Shape Data first.';
		const doc = this.pane.ownerDocument;
		this.#list.replaceChildren(
			...rows.map((row) => {
				const label = doc.createElement('label');
				const input = doc.createElement('input');
				input.type = 'checkbox';
				input.dataset.field = row.name;
				input.checked = on.has(row.name);
				input.disabled = !editable;
				const name = doc.createElement('span');
				name.textContent = row.label || row.name;
				const value = doc.createElement('small');
				value.textContent = String(row.value ?? row.rawValue ?? '');
				label.append(input, name, value);
				return label;
			}),
		);
	}
}
