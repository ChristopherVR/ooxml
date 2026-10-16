import type {
	OfficeOptionCategory,
	OfficeOptionValues,
	OfficeOptionsLabels,
} from '../../../dialog/options-dialog';
import {
	ViewerRegionHost,
	type ViewerRegionContext,
	type ViewerSettingsRegionContext,
} from '../customization/customization-regions';
import {
	getRegionRenderer,
	type ResolvedCustomization,
} from '../customization/customization-resolve';
import {
	customizeOptionsTabs,
	isOptionsPageVisible,
	isSettingLocked,
} from '../customization/customization-surfaces';
import type { OptionsSettingId } from '../customization/customization-types';
import {
	AI_OPTIONS_CATEGORY,
	buildPptxOptionCategories,
	flattenViewerOptions,
	optionValueKey,
	parseOptionValueKey,
	pptxOptionsLabels,
	type OptionsTranslate,
} from './options-dialog-model';
import { cloneViewerOptions, DEFAULT_VIEWER_OPTIONS } from './viewer-options';
import type { ViewerOptionPrimitive, ViewerOptions, ViewerOptionsGroupId } from './viewer-options';
import type { ViewerOptionsTabDefinition } from './viewer-options-controls';
import { VIEWER_OPTIONS_TABS } from './viewer-options-schema';

/** The members of `<office-ui-options-dialog>` the binding drives. */
export interface PptxOptionsDialogElement extends HTMLElement {
	categories: readonly OfficeOptionCategory[];
	values: OfficeOptionValues;
	defaults: OfficeOptionValues | null;
	labels: Partial<OfficeOptionsLabels>;
	heading: string | null;
	category: string;
	live: boolean;
	/** The host's `dialog.options` region replaces the rail, panes and footer. */
	customBody: boolean;
	show(): void;
}

/** What the dialog asks of the viewer; each binding maps these onto its own store or callbacks. */
export interface PptxOptionsDialogHost {
	/** One edit, applied at once (File > Options is live; Cancel restores the snapshot). */
	setOption(group: ViewerOptionsGroupId, key: string, value: ViewerOptionPrimitive): void;
	/** Cancel: put back the options the dialog opened with. */
	restoreOptions(options: ViewerOptions): void;
	/** Reset all options to the defaults. */
	resetOptions(): void;
	/** OK, Escape, the close button or Cancel: unmount the dialog. */
	close(): void;
	/** The selected category changed (render that category's slotted panes). */
	categoryChanged?(id: string): void;
}

/** Everything the dialog renders from; call {@link PptxOptionsDialogBinding.sync} when it changes. */
export interface PptxOptionsDialogInput {
	options: ViewerOptions;
	/** Active-language translator; a new function (or `languageKey`) re-labels the open dialog. */
	t: OptionsTranslate;
	languageKey?: string;
	customization: ResolvedCustomization;
	/** An "AI" category is offered for the chat-log export. */
	aiEnabled: boolean;
	/**
	 * The binding's region context (any `region`; `close` is ignored). When the host replaced
	 * `dialog.options`, its renderer fills the dialog body with this context plus the options
	 * model; without a context the built-in body stays.
	 */
	regionContext?: ViewerRegionContext | null;
}

export interface PptxOptionsDialogBinding {
	sync(input: PptxOptionsDialogInput): void;
	dispose(): void;
}

function lockedOptionKeys(tabs: readonly ViewerOptionsTabDefinition[]): ReadonlySet<string> {
	const keys = new Set<string>();
	for (const tab of tabs)
		for (const section of tab.sections)
			for (const control of section.controls)
				if (control.readOnly) keys.add(optionValueKey(control.group, control.key));
	return keys;
}

const DEFAULT_VALUES = flattenViewerOptions(DEFAULT_VIEWER_OPTIONS);

/**
 * Connects a `<office-ui-options-dialog live>` to a viewer: sets its categories, labels and
 * values, applies each edit through the host at once, restores the opening snapshot on Cancel and
 * resets on "Reset all options". Every binding mounts the element when File > Options opens and
 * calls this once; the panes only PowerPoint has (themes, fonts, languages, Customize Ribbon, Quick
 * Access Toolbar, Add-ins, AI) are slotted into the element by the binding.
 */
export function bindPptxOptionsDialog(
	dialog: PptxOptionsDialogElement,
	host: PptxOptionsDialogHost,
	initialCategory = 'general',
): PptxOptionsDialogBinding {
	let snapshot: ViewerOptions | null = null;
	let last: { tabs: unknown; t: OptionsTranslate; languageKey?: string; ai: boolean } | null = null;
	let opened = false;
	let lockedKeys: ReadonlySet<string> = new Set();
	let shown: OfficeOptionValues = {};
	let customization: ResolvedCustomization | null = null;
	let region: { host: ViewerRegionHost; container: HTMLElement } | null = null;
	// Whatever had focus when File > Options opened gets it back when the dialog goes away.
	const active = dialog.ownerDocument.activeElement;
	const opener =
		active instanceof HTMLElement &&
		active !== dialog.ownerDocument.body &&
		!dialog.contains(active)
			? active
			: null;

	const controller = new AbortController();
	const listen = (type: string, handler: (event: CustomEvent) => void): void =>
		dialog.addEventListener(type, handler as EventListener, { signal: controller.signal });
	listen('office-option-input', (event) => {
		const { key, value } = event.detail as { key: string; value: ViewerOptionPrimitive };
		const parsed = parseOptionValueKey(key);
		if (!parsed) return;
		// A setting the host locked never changes, whatever the control reports.
		if (lockedKeys.has(key)) dialog.values = shown;
		else host.setOption(parsed.group, parsed.key, value);
	});
	listen('office-options-reset', () => host.resetOptions());
	listen('office-options-category', (event) =>
		host.categoryChanged?.((event.detail as { id: string }).id),
	);
	listen('office-options-change', () => host.close());
	listen('office-options-cancel', () => {
		if (snapshot) host.restoreOptions(snapshot);
		host.close();
	});

	dialog.live = true;
	dialog.defaults = DEFAULT_VALUES;

	const setOption = (setting: OptionsSettingId, value: ViewerOptionPrimitive): void => {
		const parsed = parseOptionValueKey(setting);
		if (parsed && customization && !isSettingLocked(customization, setting))
			host.setOption(parsed.group, parsed.key, value);
	};
	const dropRegion = (): void => {
		region?.host.destroy();
		region?.container.remove();
		region = null;
		dialog.customBody = false;
	};
	const syncRegion = (input: PptxOptionsDialogInput): void => {
		const renderer = getRegionRenderer(input.customization, 'dialog.options');
		if (!renderer || !input.regionContext) {
			dropRegion();
			return;
		}
		if (!region) {
			const container = dialog.ownerDocument.createElement('div');
			container.slot = 'body';
			dialog.append(container);
			region = { host: new ViewerRegionHost(container), container };
		}
		dialog.customBody = true;
		const context: ViewerSettingsRegionContext = {
			...input.regionContext,
			region: 'dialog.options',
			options: input.options,
			setOption,
			resetOptions: () => host.resetOptions(),
			// The close button keeps the host's edits, as the built-in dialog's does.
			close: () => host.close(),
		};
		region.host.render(renderer, context);
	};

	return {
		sync(input) {
			snapshot ??= cloneViewerOptions(input.options);
			customization = input.customization;
			syncRegion(input);
			const tabs = customizeOptionsTabs(VIEWER_OPTIONS_TABS, input.customization);
			const ai = input.aiEnabled && isOptionsPageVisible(input.customization, AI_OPTIONS_CATEGORY);
			if (
				!last ||
				last.tabs !== tabs ||
				last.t !== input.t ||
				last.languageKey !== input.languageKey ||
				last.ai !== ai
			) {
				last = {
					tabs,
					t: input.t,
					ai,
					...(input.languageKey ? { languageKey: input.languageKey } : {}),
				};
				dialog.heading = input.t('pptx.options.title');
				dialog.labels = pptxOptionsLabels(input.t);
				dialog.categories = buildPptxOptionCategories(tabs, input.t, ai);
			}
			lockedKeys = lockedOptionKeys(tabs);
			shown = flattenViewerOptions(input.options);
			dialog.values = shown;
			if (!opened) {
				opened = true;
				dialog.category = initialCategory;
				dialog.show();
			}
		},
		dispose() {
			controller.abort();
			dropRegion();
			// The element is removed on close, so the dialog cannot hand focus back itself.
			if (opener?.isConnected) opener.focus();
		},
	};
}
