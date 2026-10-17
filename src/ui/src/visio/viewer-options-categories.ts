import {
	personalizeSection,
	unavailableCategory,
	userInterfaceSection,
	type OfficeOptionCategory,
	type OfficeOptionControl,
	type OfficeOptionSection,
	type OptionsContext,
} from '../controls';
import {
	CUSTOMIZABLE_TABS,
	DISPLAY_LANGUAGES,
	MAX_UNDO_LEVELS,
	QUICK_ACCESS_COMMANDS,
} from './viewer-behaviour';

export { MAX_UNDO_LEVELS };

const visio: OptionsContext = { product: 'Visio', t: (text) => text };
const disabled = (reason: string) => ({ disabled: `Not available. ${reason}` });
const BROWSER_SPELLING = "Your browser's spell checker decides this.";

/** Office's User Interface options, with the ones this viewer honours switched on. */
function visioUserInterface(): OfficeOptionSection {
	const section = userInterfaceSection(visio);
	return {
		...section,
		controls: section.controls.map((control): OfficeOptionControl => {
			if (control.key === 'miniToolbar') {
				const { disabled: _off, ...rest } = control;
				return {
					...rest,
					info: 'Shows text formatting beside the shape when you right-click it or select text in it.',
				};
			}
			if (control.key === 'livePreview') {
				const { disabled: _off, ...rest } = control;
				return {
					...rest,
					info: 'Hovering a style or theme in a gallery shows it on the drawing before you apply it.',
				};
			}
			if (control.key === 'screenTips')
				return {
					kind: 'select',
					key: 'screenTips',
					label: control.label,
					info: 'Applies to the ribbon. Commands that are not available always say why.',
					choices: [
						{ value: 'descriptions', label: 'Show ScreenTips with shortcut keys' },
						{ value: 'plain', label: "Don't show shortcut keys in ScreenTips" },
						{ value: 'off', label: "Don't show ScreenTips" },
					],
				};
			return control;
		}),
	};
}

const toggles = (
	prefix: 'tab' | 'qat',
	entries: readonly (readonly [string, string])[],
): OfficeOptionControl[] =>
	entries.map(([id, label]) => ({ kind: 'toggle', key: `${prefix}-${id}`, label }));

/**
 * Visio's File > Options categories, in Visio's order, built from the shared Office categories.
 * User name and Initials (the local profile, also the identity in a shared session) and Office
 * Theme (the status bar's look) are honoured, and so is everything below that says so by being
 * enabled: ScreenTips, the Mini Toolbar, Live Preview, undo levels, text auto-size, spelling, the
 * display language and which ribbon tabs and Quick Access Toolbar commands are shown. A setting the
 * viewer cannot honour is shown disabled with the reason.
 */
export const VISIO_OPTION_CATEGORIES: readonly OfficeOptionCategory[] = [
	{
		id: 'general',
		label: 'General',
		description: 'General options for working with Visio.',
		sections: [
			visioUserInterface(),
			personalizeSection(
				visio,
				[
					{ kind: 'text', key: 'userName', label: 'User name', maxLength: 64 },
					{ kind: 'text', key: 'userInitials', label: 'Initials', maxLength: 2 },
					{
						kind: 'select',
						key: 'officeTheme',
						label: 'Office Theme',
						info: 'Colorful gives the status bar the Visio colour, White keeps it neutral and Black is the dark theme.',
						choices: [
							{ value: 'colorful', label: 'Colorful' },
							{ value: 'neutral', label: 'White' },
							{ value: 'dark', label: 'Black' },
						],
					},
				],
				'Stored on this device only. Shared sessions show this name to others.',
			),
		],
	},
	{
		id: 'proofing',
		label: 'Proofing',
		description: 'Change how Visio corrects and formats your text.',
		sections: [
			{
				id: 'spelling',
				title: 'When correcting spelling in Visio',
				controls: [
					{
						kind: 'toggle',
						key: 'checkSpelling',
						label: 'Check spelling as you type',
						info: "Uses your browser's spell checker while you edit a shape's text.",
					},
				],
			},
			{
				id: 'office',
				title: 'When correcting spelling in Microsoft Office programs',
				controls: [
					'Ignore words in UPPERCASE',
					'Ignore words that contain numbers',
					'Ignore Internet and file addresses',
					'Flag repeated words',
				].map((label, index): OfficeOptionControl => ({
					kind: 'toggle',
					key: `proofing-${index}`,
					label,
					...disabled(BROWSER_SPELLING),
				})),
			},
		],
	},
	unavailableCategory(visio, 'save', {
		description: 'Customize how drawings are saved.',
		reason: 'Not available yet. Drawings are saved as downloaded copies; AutoRecover is not kept.',
	}),
	{
		id: 'language',
		label: 'Language',
		description: 'Set the Office Language Preferences.',
		sections: [
			{
				id: 'display',
				title: 'Choose Display Language',
				description:
					'The ribbon and its commands are translated. Dialogs and panes stay in English for now.',
				controls: [
					{
						kind: 'select',
						key: 'displayLanguage',
						label: 'Display language',
						info: 'Leave it on the page language to follow the page or the browser.',
						choices: [
							{ value: '', label: 'Use the page or browser language' },
							...DISPLAY_LANGUAGES.map(([value, label]) => ({ value, label })),
						],
					},
				],
			},
		],
	},
	{
		id: 'advanced',
		label: 'Advanced',
		description: 'Advanced options for working with Visio.',
		sections: [
			{
				id: 'editing',
				title: 'Editing options',
				controls: [
					{
						kind: 'number',
						key: 'undoLevels',
						label: 'Number of undo levels',
						min: 1,
						max: MAX_UNDO_LEVELS,
						info: 'Each step keeps a copy of the drawing, so very large drawings may keep fewer.',
					},
					{
						kind: 'toggle',
						key: 'autoSizeText',
						label: 'Auto-size text while editing',
						info: 'The text box grows with the text when it is taller than the shape. Off, it keeps the size of the shape and scrolls.',
					},
				],
			},
		],
	},
	{
		id: 'ribbon',
		label: 'Customize Ribbon',
		description: 'Customize the Ribbon.',
		sections: [
			{
				id: 'tabs',
				title: 'Main Tabs',
				description: 'Choose the tabs the ribbon shows. File is always shown.',
				controls: toggles('tab', CUSTOMIZABLE_TABS),
			},
		],
	},
	{
		id: 'qat',
		label: 'Quick Access Toolbar',
		description: 'Customize the Quick Access Toolbar.',
		sections: [
			{
				id: 'commands',
				title: 'Commands shown on the toolbar',
				controls: toggles('qat', QUICK_ACCESS_COMMANDS),
			},
		],
	},
	unavailableCategory(visio, 'addins'),
	unavailableCategory(visio, 'trust'),
];
