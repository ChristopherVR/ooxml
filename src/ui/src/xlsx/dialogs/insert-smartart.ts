// Excel's "Choose a SmartArt Graphic" dialog, for Insert > SmartArt and SmartArt Design > Change
// Layout: the shared `office-ui-smartart-picker` (as in Word) over the core's insert catalogue, in
// Excel's dialog frame, with every name translated through the shared SmartArt strings.
import {
	DIAGRAM_INSERT_CATEGORIES,
	DIAGRAM_INSERT_PRESETS,
	type DiagramInsertPreset,
} from 'ooxml-core/diagram';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import {
	defineSmartArtPicker,
	type OfficeSmartArtPickerLabels,
	type OfficeSmartArtValue,
} from '../../smartart';
import { showDialog } from './frame';

type Picker = HTMLElement & {
	categories: { id: string; label: string; description?: string }[];
	layouts: { id: string; category: string; label: string; family?: string }[];
	labels: Partial<OfficeSmartArtPickerLabels>;
	value: OfficeSmartArtValue;
};

/** Opens the picker; resolves with the chosen preset, or undefined when cancelled. */
export function chooseSmartArt(ctx: EditorContext): Promise<DiagramInsertPreset | undefined> {
	const doc = ctx.host.ownerDocument;
	defineSmartArtPicker(doc.defaultView?.customElements);
	const picker = doc.createElement('office-ui-smartart-picker') as Picker;
	picker.className = 'xve-smartart-picker';
	const t = (key: string) => ctx.t(key);
	picker.labels = { categories: t('SmartArt categories'), layouts: t('SmartArt layouts') };
	picker.categories = DIAGRAM_INSERT_CATEGORIES.map((category) => ({
		id: category.id,
		label: t(`smartart.category.${category.id}`),
		description: t(`smartart.description.${category.id}`),
	}));
	picker.layouts = DIAGRAM_INSERT_PRESETS.map((preset) => ({
		id: preset.layout,
		category: preset.category,
		label: t(`smartart.layout.${preset.layout}`),
		family: preset.family,
	}));
	picker.value = { layout: DIAGRAM_INSERT_PRESETS[0]?.layout ?? '', items: [] };
	const chosen = () => DIAGRAM_INSERT_PRESETS.find((entry) => entry.layout === picker.value.layout);
	return showDialog<DiagramInsertPreset>(
		ctx,
		{
			name: 'insert-smartart',
			heading: 'Choose a SmartArt Graphic',
			wide: true,
			body: picker,
			opened: () =>
				picker.shadowRoot?.querySelector<HTMLElement>('[role="option"][tabindex="0"]')?.focus(),
			submit: chosen,
		},
		(open) =>
			picker.addEventListener('office-smartart-activate', () => {
				const preset = chosen();
				if (preset) open.element.querySelector<HTMLButtonElement>('.xve-btn-primary')?.click();
			}),
	);
}
