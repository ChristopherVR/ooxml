import { INK_PEN_KINDS, MAX_INK_PENS, normalizeInkPens } from 'ooxml-core/drawingml';
import type { DrawPenPreset, PptxDrawHostOptions } from '../render/draw-canvas';
import type { RibbonDrawIntent, RibbonDrawViewState } from '../render/ribbon-draw-state';
import { readStoredViewerPrefs, writeStoredViewerPrefs } from '../render/viewer-prefs-storage';

const TYPES = INK_PEN_KINDS;
const MAX_PRESETS = MAX_INK_PENS;

/** The saved pens a deck can use: the shared gallery rules, with PowerPoint's 1 px minimum. */
export function normalizeDrawPens(value: unknown): DrawPenPreset[] {
	return normalizeInkPens(value, 16).filter((pen) => pen.width >= 1);
}

export function createDrawPens(options: {
	doc: Document;
	state: () => RibbonDrawViewState;
	host: () => PptxDrawHostOptions | undefined;
	t: (key: string) => string;
	request: (intent: RibbonDrawIntent) => void;
	activate: (kind: DrawPenPreset['kind']) => void;
}) {
	const { doc, state, host, t, request, activate } = options;
	const group = doc.createElement('pptx-ui-ribbon-group');
	group.dataset.ribbonGroup = 'draw.pens';
	const row = doc.createElement('div');
	row.className = 'advanced-row';
	const add = doc.createElement('button');
	add.type = 'button';
	add.dataset.ribbonControl = 'draw.pens.add';
	const remove = doc.createElement('button');
	remove.type = 'button';
	remove.dataset.ribbonControl = 'draw.pens.remove';
	const type = doc.createElement('pptx-ui-select');
	type.dataset.ribbonControl = 'draw.pens.type';
	const saved = doc.createElement('pptx-ui-select');
	saved.dataset.ribbonControl = 'draw.pens.presets';
	saved.setAttribute('commit-on-reselect', '');
	for (const control of [add, remove, type, saved]) {
		control.dataset.pptxCompact = '';
	}
	row.append(type, add, saved, remove);
	group.append(row);
	let presets: DrawPenPreset[] = [];
	let selectedPreset = '';
	let initialized = false;
	let store: PptxDrawHostOptions['penPreferences'];
	let previousLocale = '';
	const label = (kind: string) => t(`pptx.draw.${kind}`);
	function refresh(selected = selectedPreset) {
		const none = doc.createElement('option');
		none.value = '';
		none.textContent = t('pptx.draw.presets');
		saved.replaceChildren(
			none,
			...presets.map((preset, index) => {
				const item = doc.createElement('option');
				item.value = String(index);
				item.textContent = `${label(preset.kind)} ${preset.color} (${preset.width} px)`;
				return item;
			}),
		);
		selectedPreset = selected && Number(selected) < presets.length ? selected : '';
		saved.value = selectedPreset;
		remove.disabled = !state().editable || !selectedPreset;
	}
	function persist() {
		try {
			if (store) {
				store.write(presets);
			} else {
				writeStoredViewerPrefs({ inkPens: presets });
			}
		} catch {
			// Retain the in-memory preset when the host storage is unavailable.
		}
	}
	function apply(preset: DrawPenPreset) {
		type.value = preset.kind;
		request({ kind: 'preset', value: preset });
		activate(preset.kind);
	}
	add.addEventListener('click', () => {
		const { color, width, editable } = state();
		const kind = TYPES.find((item) => item === type.value);
		if (!kind || !editable || !/^#[\da-f]{6}$/iu.test(color) || width < 1 || width > 16) {
			return;
		}
		const preset = { kind, color, width };
		const existing = presets.findIndex(
			(pen) => pen.kind === kind && pen.color === color && pen.width === width,
		);
		if (existing < 0 && presets.length < MAX_PRESETS) {
			presets = [...presets, preset];
			persist();
		}
		refresh(String(existing < 0 ? presets.length - 1 : existing));
		apply(preset);
		add.disabled = presets.length >= MAX_PRESETS;
	});
	remove.addEventListener('click', () => {
		if (!remove.disabled) {
			presets = presets.filter((_, index) => index !== Number(selectedPreset));
			persist();
			refresh('');
			add.disabled = false;
			saved.focus();
		}
	});
	saved.addEventListener('change', () => {
		selectedPreset = saved.value;
		const preset = selectedPreset === '' ? undefined : presets[Number(selectedPreset)];
		if (preset && state().editable) {
			apply(preset);
		}
		remove.disabled = !preset || !state().editable;
	});
	return {
		group,
		sync() {
			const nextStore = host()?.penPreferences;
			const locale = TYPES.map(label).join('|') + t('pptx.draw.presets');
			if (!initialized || store !== nextStore || previousLocale !== locale) {
				if (!initialized || store !== nextStore) {
					store = nextStore;
					try {
						presets = normalizeDrawPens(store ? store.read() : readStoredViewerPrefs().inkPens);
					} catch {
						presets = [];
					}
				}
				initialized = true;
				previousLocale = locale;
				type.replaceChildren(
					...TYPES.map((kind) => {
						const option = doc.createElement('option');
						option.value = kind;
						option.textContent = label(kind);
						return option;
					}),
				);
				refresh();
			}
			group.setAttribute('label', t('pptx.draw.presets'));
			add.textContent = t('pptx.draw.addPen');
			remove.textContent = t('pptx.draw.removePen');
			type.setAttribute('aria-label', t('pptx.draw.penType'));
			saved.setAttribute('aria-label', t('pptx.draw.presets'));
			type.disabled = saved.disabled = !state().editable;
			add.disabled = !state().editable || presets.length >= MAX_PRESETS;
			add.title = presets.length >= MAX_PRESETS ? t('pptx.draw.penLimit') : '';
			remove.disabled = !state().editable || !selectedPreset;
		},
	};
}
