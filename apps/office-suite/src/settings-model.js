/**
 * The suite's Settings as data for the shared `office-ui-options-dialog`, in the arrangement of
 * PowerPoint's File > Options: a category rail, a headline, small section titles and rows. Pure,
 * so it is tested without a DOM. The panes the rows cannot express (appearance cards, fonts,
 * privacy, assistant) are slots that `settings-dialog.js` fills.
 */
import { CORNERS, DEFAULT_SETTINGS, LOCALES } from './settings.js';

/** Slot of a block under a section's rows. */
export const blockSlot = (name) => `block-${name}`;

const toggle = (key, label, info) => ({ kind: 'toggle', key, label, ...(info ? { info } : {}) });

/**
 * The categories. `profileName` fills the user name hint; `assistant` adds the Assistant
 * category when the deployment's policy turns the assistant on.
 */
export function settingsCategories({ profileName = '', assistant = false } = {}) {
	const categories = [
		{
			id: 'general',
			label: 'General',
			description: 'General options for working with OOXML Office.',
			sections: [
				{
					id: 'personalize',
					title: 'Personalize your copy of OOXML Office',
					controls: [
						{
							kind: 'text',
							key: 'authorName',
							label: 'User name',
							maxLength: 80,
							info: `The author of the comments and tracked changes you add in Word, Excel and PowerPoint. Leave it blank to use your profile name${profileName ? ` (${profileName})` : ''}.`,
						},
					],
				},
				{
					id: 'appearance',
					title: 'Appearance',
					description:
						'Colours the workspace and every open app. Documents keep their own colours.',
					controls: [
						{
							kind: 'select',
							key: 'corners',
							label: 'Corner style',
							choices: Object.entries(CORNERS).map(([value, { label }]) => ({ value, label })),
						},
						toggle(
							'reduceMotion',
							'Reduce motion',
							'Turns off the workspace transitions and the theme cross-fade. Your device setting for reduced motion is always honoured.',
						),
					],
					slot: blockSlot('appearance'),
				},
				{
					id: 'closing',
					title: 'Leaving the app',
					controls: [toggle('confirmClose', 'Warn me before leaving with unsaved changes')],
				},
			],
		},
		{
			id: 'view',
			label: 'View',
			description: 'Choose what the apps show around your documents.',
			sections: [
				{
					id: 'chrome',
					title: 'Show',
					description:
						'Open documents change at once where the app allows it; the rest follow the next time you open a document.',
					controls: [
						toggle('showRibbon', 'Ribbon', 'Word, Excel, PowerPoint and Visio.'),
						toggle('showThumbnails', 'Page and slide thumbnails', 'Word and PowerPoint.'),
						toggle('showFormulaBar', 'Formula bar', 'Excel.'),
					],
				},
				{
					id: 'units',
					title: 'Measurement',
					controls: [
						{
							kind: 'select',
							key: 'measurementUnit',
							label: 'Ruler units',
							info: 'PowerPoint rulers and guides.',
							choices: [
								{ value: 'inches', label: 'Inches' },
								{ value: 'centimetres', label: 'Centimetres' },
							],
						},
					],
				},
			],
		},
		{
			id: 'language',
			label: 'Language',
			description: 'Set the language of the apps.',
			sections: [
				{
					id: 'display',
					title: 'Display language',
					description:
						'Menus, ribbons, dialogs and messages in Word, Excel, PowerPoint and Visio. The workspace around them is in English for now.',
					controls: [
						{
							kind: 'select',
							key: 'locale',
							label: 'Display language',
							choices: [
								{ value: 'auto', label: "Each app's own choice" },
								...LOCALES.map(([value, label]) => ({ value, label })),
							],
						},
					],
				},
			],
		},
		{
			id: 'fonts',
			label: 'Fonts',
			description: 'How documents get the fonts they use but this device does not have.',
			sections: [],
			custom: true,
		},
		{
			id: 'privacy',
			label: 'Privacy',
			description: 'What OOXML Office connects to, and what it keeps on this device.',
			sections: [
				{
					id: 'linked',
					title: 'Linked content',
					description:
						'Presentations can link pictures, video and audio from other websites instead of embedding them. Loading them tells those sites that the file was opened.',
					controls: [
						toggle(
							'blockExternalContent',
							'Block linked pictures and media in presentations',
							"Off leaves the choice to PowerPoint's Trust Center (on by default there). On blocks it in every presentation you open next and locks that option.",
						),
					],
				},
			],
			custom: true,
		},
	];
	if (assistant)
		categories.push({
			id: 'assistant',
			label: 'Assistant',
			description: 'The AI assistant beside your documents.',
			sections: [],
			custom: true,
		});
	return categories;
}

/** The dialog's values record: every setting is a primitive already. */
export const settingsValues = (current) => ({ ...current });

/** "Reset all options" puts these back (appearance and fonts are reset beside them). */
export const settingsDefaults = () => ({ ...DEFAULT_SETTINGS });

export const settingsLabels = {
	ok: 'OK',
	cancel: 'Cancel',
	reset: 'Reset all settings',
	categories: 'Settings categories',
	close: 'Close settings',
};
