import {
	ACTION_BUTTON_PRESETS,
	actionButtonLabelKey,
	FREEFORM_TOOL_IDS,
	FREEFORM_TOOL_LABEL_KEYS,
	INSERT_CHART_TYPES,
	INSERT_GROUPS,
	INSERT_GROUP_CAPTIONS,
	SHAPE_PRESET_DEFS,
	insertControlStatus,
	insertEntryEnabled,
	insertLabel,
} from '../render';
import type {
	InsertControlSpec,
	InsertLabel,
	RibbonInsertIntent,
	RibbonInsertState,
} from '../render';
import { createInsertMenu } from './ribbon-insert-dom';
import type { InsertMenu, InsertMenuItem } from './ribbon-insert-dom';
import {
	INSERT_SHAPE_GLYPH_PATHS,
	RIBBON_INSERT_ICON_PATHS,
	insertGlyphTransform,
} from './ribbon-insert-icons';

const FREEFORM_ICONS: Record<string, string> = { freeformShape: 'penTool', curve: 'spline' };
const FREEFORM_FALLBACKS: Record<string, string> = {
	freeformShape: 'Freeform: Shape',
	curve: 'Curve',
};
const UNSUPPORTED: InsertLabel = ['pptx.insert.unavailable', 'Not available in this viewer yet.'];
const NEEDS_SELECTION: InsertLabel = ['pptx.insert.needsSelection', 'Select an object first.'];

interface BuiltControl {
	spec: InsertControlSpec;
	el: HTMLElement;
	sync(state: RibbonInsertState, text: (label: InsertLabel) => string): void;
}

/** All Insert markup lives here. The host only reflects state and routes native intents. */
export function createRibbonInsertView(
	doc: Document,
	request: (intent: RibbonInsertIntent) => void,
) {
	const groups = new Map<string, HTMLElement>();
	const group = (id: string, ...children: HTMLElement[]) => {
		const el = doc.createElement('pptx-ui-ribbon-group');
		el.dataset.ribbonGroup = id;
		el.append(...children);
		groups.set(id, el);
		return el;
	};
	const stack = (...children: HTMLElement[]) => {
		const el = doc.createElement('div');
		el.className = 'stack';
		el.append(...children);
		return el;
	};
	const command = (icon: string, compact: boolean, run: () => void, control?: string) => {
		const el = doc.createElement('pptx-ui-ribbon-command');
		if (control) {
			el.dataset.ribbonControl = control;
		}
		el.setAttribute('icon', icon);
		el.toggleAttribute('compact', compact);
		// Controls without a catalogue id cannot emit command-request, so use the click.
		el.addEventListener('command-request', (event) => event.stopPropagation());
		el.addEventListener('click', () => {
			if (!el.hasAttribute('disabled')) {
				run();
			}
		});
		return el;
	};
	/**
	 * A gallery pick stages the type and then inserts it, as Office does. The host applies the
	 * staged type on its next render, so the insert follows after it has settled.
	 */
	const stageThenInsert = (stage: RibbonInsertIntent, insert: RibbonInsertIntent) => {
		request(stage);
		setTimeout(() => request(insert), 50);
	};
	const menus: InsertMenu[] = [];
	const built: BuiltControl[] = [];
	/** ScreenTip: the title, its description and, when disabled, the honest reason. */
	const tooltip = (
		spec: InsertControlSpec,
		text: (label: InsertLabel) => string,
		reason: string | null,
	) => {
		const lines = [text(spec.label), text(spec.tip)];
		if (reason === 'unsupported') {
			lines.push(text(UNSUPPORTED));
		} else if (reason === 'selection') {
			lines.push(text(NEEDS_SELECTION));
		}
		return lines.join('\n');
	};
	const addMenu = (spec: InsertControlSpec, menu: InsertMenu, sync: BuiltControl['sync']) => {
		menus.push(menu);
		built.push({ spec, el: menu.el, sync });
	};

	const buildShapes = (spec: InsertControlSpec) => {
		const menu = createInsertMenu(
			doc,
			spec.control,
			'shapes',
			(value) => stageThenInsert({ kind: 'shapeType', value }, { kind: 'shape', value }),
			{ large: true, grid: true },
		);
		addMenu(spec, menu, (state, text) => {
			menu.sync({
				label: text(spec.label),
				title: tooltip(spec, text, null),
				items: SHAPE_PRESET_DEFS.map((item) => ({
					id: item.type,
					label: text([item.i18nKey, item.label]),
					glyph: INSERT_SHAPE_GLYPH_PATHS[item.glyph],
					viewBox: '0 0 16 16',
					transform: insertGlyphTransform(item.glyphClass),
				})),
				disabled: !state.editable,
				menuLabel: text(['pptx.insert.shapeType', 'Shape type']),
			});
		});
		return menu.el;
	};
	const buildChart = (spec: InsertControlSpec) => {
		const menu = createInsertMenu(
			doc,
			spec.control,
			'chart',
			(value) => stageThenInsert({ kind: 'chartType', value }, { kind: 'chart', value }),
			{ large: true },
		);
		addMenu(spec, menu, (state, text) => {
			menu.sync({
				label: text(spec.label),
				title: tooltip(spec, text, null),
				items: INSERT_CHART_TYPES.map((item) => ({
					id: item.id,
					label: text([item.labelKey, item.label]),
					glyph: RIBBON_INSERT_ICON_PATHS.chart,
					viewBox: '0 0 20 20',
				})),
				disabled: !state.editable,
				menuLabel: text(['pptx.ribbon.chartType', 'Chart type']),
			});
		});
		return menu.el;
	};
	const buildAction = (spec: InsertControlSpec) => {
		const menu = createInsertMenu(
			doc,
			spec.control,
			'action',
			(value) => request({ kind: 'actionButton', value }),
			{ large: true },
		);
		addMenu(spec, menu, (state, text) => {
			const actionTitle = text(['pptx.ribbon.insertActionButton', 'Insert action button']);
			menu.sync({
				label: text(spec.label),
				title: tooltip(spec, text, null),
				items: ACTION_BUTTON_PRESETS.map((item) => ({
					id: item.shapeType,
					label: text([actionButtonLabelKey(item.shapeType), item.label]),
					glyph: item.iconPath,
				})),
				disabled: !state.editable,
				menuLabel: actionTitle,
			});
		});
		return menu.el;
	};
	/** New Slide: the face adds a slide, the chevron offers the host's layouts. */
	const buildLayouts = (spec: InsertControlSpec) => {
		const menu = createInsertMenu(
			doc,
			spec.control,
			spec.icon,
			(value) => request({ kind: 'newSlide', value }),
			{ large: true, split: { primary: () => request({ kind: 'newSlide', value: null }) } },
		);
		addMenu(spec, menu, (state, text) => {
			const status = insertControlStatus(state, spec);
			const layouts = state.layouts ?? [];
			menu.sync({
				label: text(spec.label),
				title: tooltip(spec, text, status.reason),
				menuTitle: `${text(spec.label)}: ${text(['pptx.ribbon.groupLayouts', 'Layouts'])}`,
				items: layouts.map((item) => ({
					id: item.path,
					label: item.name,
					section: text(['pptx.ribbon.groupLayouts', 'Layouts']),
					current: item.path === state.currentLayoutPath,
				})),
				disabled: status.disabled || layouts.length === 0,
				primaryDisabled: status.disabled,
				menuLabel: text(['pptx.ribbon.groupLayouts', 'Layouts']),
			});
		});
		return menu.el;
	};
	/** Any other dropdown: its entries are listed, those without a backing feature disabled. */
	const buildMenu = (spec: InsertControlSpec) => {
		const entries = spec.menu ?? [];
		const menu = createInsertMenu(
			doc,
			spec.control,
			spec.icon,
			(id) => {
				const item = entries.find((candidate) => candidate.id === id);
				if (item?.intent) {
					request(item.intent);
				}
			},
			{
				large: spec.large,
				...(spec.split && spec.intent ? { split: { primary: () => request(spec.intent!) } } : {}),
			},
		);
		addMenu(spec, menu, (state, text) => {
			const status = insertControlStatus(state, spec);
			const items: InsertMenuItem[] = entries.map((item) => {
				const enabled = insertEntryEnabled(state, item);
				return {
					id: item.id,
					label: text(item.label),
					disabled: !enabled,
					...(item.intent ? {} : { title: text(UNSUPPORTED) }),
				};
			});
			menu.sync({
				label: text(spec.label),
				title: tooltip(spec, text, status.reason),
				menuTitle: `${text(spec.label)}: ${text(['pptx.insert.options', 'Options'])}`,
				items,
				disabled: status.disabled,
				menuLabel: spec.split ? text(['pptx.insert.options', 'Options']) : text(spec.label),
			});
		});
		return menu.el;
	};
	/** A plain button: large or small, enabled only when the editor can act on it. */
	const buildButton = (spec: InsertControlSpec) => {
		const el = command(
			spec.icon,
			!spec.large,
			() => spec.intent && request(spec.intent),
			spec.control,
		);
		built.push({
			spec,
			el,
			sync(state, text) {
				const status = insertControlStatus(state, spec);
				el.setAttribute('label', text(spec.label));
				el.setAttribute('title', tooltip(spec, text, status.reason));
				el.toggleAttribute('disabled', status.disabled);
				el.hidden = status.hidden;
			},
		});
		return el;
	};
	const build = (spec: InsertControlSpec) => {
		switch (spec.special) {
			case 'shapes':
				return buildShapes(spec);
			case 'chart':
				return buildChart(spec);
			case 'action':
				return buildAction(spec);
			case 'layouts':
				return buildLayouts(spec);
			default:
				return spec.menu ? buildMenu(spec) : buildButton(spec);
		}
	};

	let armed: string | null = null;
	const freeform = FREEFORM_TOOL_IDS.map((tool) => {
		const el = command(FREEFORM_ICONS[tool], true, () =>
			request({ kind: 'freeform', value: armed === tool ? null : tool }),
		);
		el.dataset.pptxDrawingTool = tool;
		return { tool, el };
	});
	const freeformStack = stack(...freeform.map(({ el }) => el));

	const root = INSERT_GROUPS.map((spec) => {
		const children: HTMLElement[] = [];
		for (const item of spec.items) {
			if (Array.isArray(item)) {
				children.push(stack(...(item as readonly InsertControlSpec[]).map(build)));
			} else {
				const control = item as InsertControlSpec;
				children.push(build(control));
				if (control.special === 'shapes') {
					children.push(freeformStack);
				}
			}
		}
		return group(spec.id, ...children);
	});
	const sync = (state: RibbonInsertState) => {
		const text = (label: InsertLabel) => insertLabel(state, label);
		armed = state.activeFreeformTool ?? null;
		const disabled = !state.editable;
		for (const [id, el] of groups) {
			el.setAttribute('label', text(INSERT_GROUP_CAPTIONS[id]));
		}
		for (const control of built) {
			control.sync(state, text);
			control.el.hidden = insertControlStatus(state, control.spec).hidden;
		}
		const visible = state.freeformTools ?? [];
		freeformStack.hidden = visible.length === 0;
		for (const { tool, el } of freeform) {
			const label = text([FREEFORM_TOOL_LABEL_KEYS[tool], FREEFORM_FALLBACKS[tool]]);
			el.setAttribute('label', label);
			el.setAttribute('title', label);
			el.hidden = !visible.includes(tool);
			el.setAttribute('pressed', String(armed === tool));
			el.toggleAttribute('active', armed === tool);
			el.toggleAttribute('disabled', disabled);
		}
	};
	return { root, sync, menus };
}
