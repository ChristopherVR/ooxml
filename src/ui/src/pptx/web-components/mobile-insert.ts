import {
	INSERT_CHART_TYPES,
	INSERT_GROUPS,
	SHAPE_PRESET_DEFS,
	canRequestInsert,
	insertControlStatus,
	insertLabel,
} from '../render';
import type {
	InsertControlSpec,
	InsertLabel,
	RibbonInsertIntent,
	RibbonInsertState,
} from '../render';
import { attachControlStyles } from './control-styles';
import { MOBILE_INSERT_STYLES } from './mobile-insert-styles';
import { insertIcon, insertSvg } from './ribbon-insert-dom';
import { INSERT_SHAPE_GLYPH_PATHS, insertGlyphTransform } from './ribbon-insert-icons';
import { createTablePicker } from './ribbon-insert-table-picker';

export type MobileInsertRequestEvent = CustomEvent<RibbonInsertIntent>;
export interface PptxUiMobileInsertElement extends HTMLElement {
	state: RibbonInsertState;
	/** Back to the tile grid, e.g. when the host reopens the sheet. */
	reset(): void;
}
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-mobile-insert': PptxUiMobileInsertElement;
	}
}

/** The phone sheet's tiles, in PowerPoint mobile's order, by the shared control id. */
const TILES = [
	'insert.text.textBox',
	'insert.illustrations.shapes',
	'insert.images.pictures',
	'insert.tables.table',
	'insert.illustrations.chart',
	'insert.illustrations.smartArt',
	'insert.symbols.equation',
	'insert.media.media',
	'insert.slides.newSlide',
	'insert.comments.comment',
] as const;

type View = 'home' | 'shapes' | 'table' | 'chart' | 'layouts';

const SPECS = new Map<string, InsertControlSpec>(
	(INSERT_GROUPS.flatMap((group) => group.items.flat()) as InsertControlSpec[]).map((spec) => [
		spec.control,
		spec,
	]),
);
const CHEVRON = 'M12 4 6 10l6 6';
const MORE = 'm5 8 5 5 5-5';

/**
 * The phone Insert sheet body: PowerPoint mobile's tile grid (Text Box, Shapes, Pictures, Table,
 * Chart, SmartArt, Equation, Video, New Slide, Comment). It draws the same shared Insert
 * definitions as `pptx-ui-ribbon-insert` (labels, icons, availability, the shape, chart and
 * layout catalogues and the Table grid) and emits the same `insert-request` intents, so a host
 * routes it exactly like the ribbon tab and adds no command logic of its own.
 */
export function definePptxMobileInsert(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-mobile-insert')) {
		return;
	}
	class MobileInsert extends HTMLElement implements PptxUiMobileInsertElement {
		private model: RibbonInsertState = {
			editable: false,
			hasSelection: false,
			shapeType: 'rect',
			chartKind: 'column',
		};
		private view: View = 'home';
		private readonly body = this.ownerDocument.createElement('div');
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachControlStyles(root, MOBILE_INSERT_STYLES);
			root.append(this.body);
		}
		get state() {
			return this.model;
		}
		set state(value: RibbonInsertState) {
			this.model = value;
			this.render();
		}
		connectedCallback(): void {
			this.render();
		}
		reset(): void {
			this.view = 'home';
			this.render();
		}
		private text(label: InsertLabel): string {
			return insertLabel(this.model, label);
		}
		private request(intent: RibbonInsertIntent): void {
			if (canRequestInsert(this.model, intent)) {
				this.dispatchEvent(
					new CustomEvent('insert-request', { detail: intent, bubbles: true, composed: true }),
				);
			}
		}
		private open(view: View): void {
			this.view = view;
			this.render();
		}
		private button(className: string, label: string, ...children: Node[]): HTMLButtonElement {
			const button = this.ownerDocument.createElement('button');
			button.type = 'button';
			button.className = className;
			button.setAttribute('aria-label', label);
			button.append(...children);
			return button;
		}
		private tile(spec: InsertControlSpec): HTMLButtonElement | null {
			const status = insertControlStatus(this.model, spec);
			if (status.hidden) {
				return null;
			}
			const doc = this.ownerDocument;
			const label = this.text(spec.label);
			const submenu: View | null =
				spec.special === 'shapes'
					? 'shapes'
					: spec.special === 'chart'
						? 'chart'
						: spec.picker === 'table'
							? 'table'
							: spec.special === 'layouts' && (this.model.layouts ?? []).length > 0
								? 'layouts'
								: null;
			const caption = doc.createElement('span');
			caption.textContent = label;
			const tile = this.button('tile', label, insertIcon(doc, spec.icon), caption);
			if (submenu) {
				const more = insertSvg(doc, MORE, '0 0 20 20');
				more.classList.add('more');
				tile.append(more);
			}
			tile.dataset.insertControl = spec.control;
			const intent = spec.intent ?? spec.menu?.find((item) => item.intent)?.intent;
			tile.disabled = submenu ? status.disabled : status.disabled || !intent;
			tile.title = this.text(spec.tip);
			tile.addEventListener('click', () => {
				if (submenu) {
					this.open(submenu);
				} else if (intent) {
					this.request(intent);
				}
			});
			return tile;
		}
		private head(spec: InsertControlSpec): HTMLElement {
			const doc = this.ownerDocument;
			const head = doc.createElement('div');
			head.className = 'head';
			const back = this.button(
				'back',
				this.text(['pptx.ribbon.tab.insert', 'Insert']),
				insertSvg(doc, CHEVRON, '0 0 20 20'),
				doc.createTextNode(this.text(['pptx.ribbon.tab.insert', 'Insert'])),
			);
			back.addEventListener('click', () => this.open('home'));
			const title = doc.createElement('strong');
			title.textContent = this.text(spec.label);
			head.append(back, title);
			return head;
		}
		private render(): void {
			const doc = this.ownerDocument;
			const children: HTMLElement[] = [];
			if (this.view === 'home') {
				const grid = doc.createElement('div');
				grid.className = 'grid';
				for (const id of TILES) {
					const tile = this.tile(SPECS.get(id) as InsertControlSpec);
					if (tile) {
						grid.append(tile);
					}
				}
				children.push(grid);
			} else if (this.view === 'shapes') {
				children.push(this.head(SPECS.get('insert.illustrations.shapes') as InsertControlSpec));
				const shapes = doc.createElement('div');
				shapes.className = 'shapes';
				for (const item of SHAPE_PRESET_DEFS) {
					const label = this.text([item.i18nKey, item.label]);
					const glyph = insertSvg(doc, INSERT_SHAPE_GLYPH_PATHS[item.glyph] ?? '', '0 0 16 16');
					const transform = insertGlyphTransform(item.glyphClass);
					if (transform !== 'none') {
						glyph.style.transform = transform;
					}
					const button = this.button('shape', label, glyph);
					button.title = label;
					button.disabled = !this.model.editable;
					button.addEventListener('click', () => this.request({ kind: 'shape', value: item.type }));
					shapes.append(button);
				}
				children.push(shapes);
			} else if (this.view === 'chart') {
				children.push(this.head(SPECS.get('insert.illustrations.chart') as InsertControlSpec));
				const rows = doc.createElement('div');
				rows.className = 'rows';
				for (const item of INSERT_CHART_TYPES) {
					const label = this.text([item.labelKey, item.label]);
					const row = this.button(
						'row',
						label,
						insertIcon(doc, 'chart'),
						doc.createTextNode(label),
					);
					row.addEventListener('click', () => this.request({ kind: 'chart', value: item.id }));
					rows.append(row);
				}
				children.push(rows);
			} else if (this.view === 'layouts') {
				children.push(this.head(SPECS.get('insert.slides.newSlide') as InsertControlSpec));
				const rows = doc.createElement('div');
				rows.className = 'rows';
				for (const item of this.model.layouts ?? []) {
					const row = this.button('row', item.name, doc.createTextNode(item.name));
					row.toggleAttribute('aria-current', item.path === this.model.currentLayoutPath);
					row.addEventListener('click', () => this.request({ kind: 'newSlide', value: item.path }));
					rows.append(row);
				}
				children.push(rows);
			} else {
				children.push(this.head(SPECS.get('insert.tables.table') as InsertControlSpec));
				const picker = createTablePicker(doc, (rows, columns) =>
					this.request({ kind: 'table', rows, columns }),
				);
				picker.disabled = !this.model.editable;
				picker.sync({
					heading: this.text(['pptx.insert.tablePicker', 'Insert Table']),
					size: (columns, rows) =>
						this.text(['pptx.insert.tableSize', '{{columns}}x{{rows}} Table'])
							.replace('{{columns}}', String(columns))
							.replace('{{rows}}', String(rows)),
				});
				children.push(picker.el);
			}
			this.body.replaceChildren(...children);
		}
	}
	registry.define('pptx-ui-mobile-insert', MobileInsert);
}
