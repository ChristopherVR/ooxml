// Word's "Choose a SmartArt Graphic" dialog, for Insert > SmartArt and SmartArt Design > Change
// Layout: the shared `office-ui-smartart-picker` (categories, layout gallery, preview) in the
// editor's dialog frame, over the core's insert catalogue, with every name translated.
import {
	DIAGRAM_INSERT_CATEGORIES,
	DIAGRAM_INSERT_PRESETS,
	type DiagramInsertPreset,
} from 'ooxml-core/diagram';
import {
	defineSmartArtPicker,
	type OfficeSmartArtPickerLabels,
	type OfficeSmartArtValue,
} from '../smartart';
import { dialogButton } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import { translate, type EditorLocale, type LocalizationKey } from './localization';

type Picker = HTMLElement & {
	categories: { id: string; label: string; description?: string }[];
	layouts: { id: string; category: string; label: string; family?: string }[];
	labels: Partial<OfficeSmartArtPickerLabels>;
	value: OfficeSmartArtValue;
};

export interface SmartArtDialog {
	readonly element: HTMLElement;
	readonly isOpen: boolean;
	/** Opens the dialog; `current` preselects a layout (Change Layout). */
	open(choose: (preset: DiagramInsertPreset) => void, current?: string): void;
	close(): void;
	setLocale(locale: EditorLocale): void;
}

/** The dialog; `restoreFocus` runs when it closes. */
export function createSmartArtDialog(restoreFocus: () => void): SmartArtDialog {
	defineSmartArtPicker();
	const element = createDialogShell('Choose a SmartArt Graphic', 'dve-smartart-dialog');
	const picker = document.createElement('office-ui-smartart-picker') as Picker;
	picker.className = 'dve-smartart-picker';
	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('OK', true);
	element.append(picker, dialogActions(cancel, confirm));
	let locale: EditorLocale = 'en';
	let onChoose: ((preset: DiagramInsertPreset) => void) | undefined;

	const t = (key: string) => translate(locale, key as LocalizationKey);
	const fill = () => {
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
	};
	const hide = () => {
		setDialogOpen(element, false);
		restoreFocus();
	};
	const submit = () => {
		const preset = DIAGRAM_INSERT_PRESETS.find((entry) => entry.layout === picker.value.layout);
		const choose = onChoose;
		hide();
		if (preset && choose) choose(preset);
	};
	cancel.addEventListener('click', hide);
	confirm.addEventListener('click', submit);
	picker.addEventListener('office-smartart-activate', submit);
	onDialogDismiss(element, hide);
	return {
		element,
		get isOpen() {
			return isDialogOpen(element);
		},
		open(choose, current) {
			onChoose = choose;
			fill();
			localizeDialog(element, locale);
			const first = DIAGRAM_INSERT_PRESETS[0]?.layout ?? '';
			picker.value = { layout: current ?? first, items: [] };
			setDialogOpen(element, true);
			picker.shadowRoot?.querySelector<HTMLElement>('[role="option"][tabindex="0"]')?.focus();
		},
		close: hide,
		setLocale(next) {
			locale = next;
		},
	};
}
