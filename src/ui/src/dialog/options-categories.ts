import type {
	OfficeOptionCategory,
	OfficeOptionControl,
	OfficeOptionSection,
	OfficeOptionsLabels,
} from './options-dialog';

/**
 * The File > Options categories every product shares, as builders over the product's translate
 * function (English keys, see `optionStrings` in `ooxml-core/i18n`). A product supplies the
 * settings it really honours and takes the rest from here: the Office category order, and the
 * honest "not available yet" panes for what it cannot do, so Word, Excel and Visio read like
 * PowerPoint's Options. Nothing here pretends to work.
 */
export interface OptionsContext {
	/** Product name for the `{product}` placeholders: "Word", "Excel", "Visio". */
	product: string;
	/** Display-language lookup; identity for English-only products. */
	t(text: string): string;
}

export type UnavailableKind =
	| 'proofing'
	| 'save'
	| 'language'
	| 'accessibility'
	| 'advanced'
	| 'ribbon'
	| 'qat'
	| 'addins'
	| 'trust';

interface UnavailableText {
	id: string;
	label: string;
	description: string;
	reason: string;
}

const UNAVAILABLE: Record<UnavailableKind, UnavailableText> = {
	proofing: {
		id: 'proofing',
		label: 'Proofing',
		description: 'Change how {product} corrects and formats your text.',
		reason: 'Not available yet. Needs rich text formatting.',
	},
	save: {
		id: 'save',
		label: 'Save',
		description: 'Customize how documents are saved.',
		reason: 'Not available yet. Documents are saved to this device; AutoRecover is not kept.',
	},
	language: {
		id: 'language',
		label: 'Language',
		description: 'Set the Office Language Preferences.',
		reason: 'Not available yet. The display language is fixed.',
	},
	accessibility: {
		id: 'accessibility',
		label: 'Accessibility',
		description: 'Make {product} easier to work with.',
		reason: 'Not available yet. There are no accessibility settings to change.',
	},
	advanced: {
		id: 'advanced',
		label: 'Advanced',
		description: 'Advanced options for working with {product}.',
		reason: 'Not available yet. The editing behaviour is fixed.',
	},
	ribbon: {
		id: 'ribbon',
		label: 'Customize Ribbon',
		description: 'Customize the Ribbon and keyboard shortcuts.',
		reason: 'Not available yet. The ribbon layout is fixed.',
	},
	qat: {
		id: 'qat',
		label: 'Quick Access Toolbar',
		description: 'Customize the Quick Access Toolbar.',
		reason: 'Not available yet. The Quick Access Toolbar is fixed.',
	},
	addins: {
		id: 'add-ins',
		label: 'Add-ins',
		description: 'View and manage Microsoft Office Add-ins.',
		reason: 'Not available. Add-ins and macros are never run.',
	},
	trust: {
		id: 'trust',
		label: 'Trust Center',
		description: 'Help keep your documents safe and your computer secure and healthy.',
		reason: 'Not available. Documents stay local and active content is never run.',
	},
};

/** A pane's English text with `{product}` filled in, translated first. */
export function optionText(context: OptionsContext, text: string): string {
	return context.t(text).replace('{product}', context.product);
}

/**
 * A category the product cannot offer yet: it keeps its place in the list and says why. `text`
 * replaces the shared wording where a product has something more specific to say (its text goes
 * through `t` like the rest).
 */
export function unavailableCategory(
	context: OptionsContext,
	kind: UnavailableKind,
	text: Partial<Pick<UnavailableText, 'description' | 'reason'>> = {},
): OfficeOptionCategory {
	const entry = UNAVAILABLE[kind];
	return {
		id: entry.id,
		label: context.t(entry.label),
		description: optionText(context, text.description ?? entry.description),
		sections: [],
		disabled: optionText(context, text.reason ?? entry.reason),
	};
}

/** A disabled row for something Office offers that this product does not do yet. */
export function unavailableControl(
	context: OptionsContext,
	control: OfficeOptionControl,
	reason: string,
): OfficeOptionControl {
	return { ...control, disabled: context.t(reason) };
}

/** "User Interface options": Office's three display settings, none of which are drawn yet. */
export function userInterfaceSection(context: OptionsContext): OfficeOptionSection {
	const { t } = context;
	return {
		id: 'ui',
		title: t('User Interface options'),
		controls: [
			{
				kind: 'toggle',
				key: 'miniToolbar',
				label: t('Show Mini Toolbar on selection'),
				disabled: t('Not available yet. Needs rich text formatting.'),
			},
			{
				kind: 'toggle',
				key: 'livePreview',
				label: t('Enable Live Preview'),
				disabled: t(
					'Not available yet. Styles are applied on click; hover previews are not drawn yet.',
				),
			},
			{
				kind: 'select',
				key: 'screenTips',
				label: t('ScreenTip style'),
				choices: [{ value: 'descriptions', label: t('Show feature descriptions in ScreenTips') }],
				disabled: t('Not available yet. ScreenTips always describe the command.'),
			},
		],
	};
}

/** "Personalize your copy of Microsoft Office": the product's own name, theme and the like. */
export function personalizeSection(
	context: OptionsContext,
	controls: readonly OfficeOptionControl[],
	description?: string,
): OfficeOptionSection {
	return {
		id: 'personalize',
		title: context.t('Personalize your copy of Microsoft Office'),
		...(description ? { description } : {}),
		controls,
	};
}

/** The dialog's own buttons in the product's language. */
export function optionsLabels(context: OptionsContext): OfficeOptionsLabels {
	const { t } = context;
	return {
		ok: t('OK'),
		cancel: t('Cancel'),
		reset: t('Reset all options'),
		categories: t('Categories'),
		close: t('Close'),
	};
}
