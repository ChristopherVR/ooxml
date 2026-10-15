import {
	personalizeSection,
	unavailableCategory,
	userInterfaceSection,
	type OfficeOptionCategory,
	type OptionsContext,
} from '../controls';

const visio: OptionsContext = { product: 'Visio', t: (text) => text };
const disabled = (reason: string) => ({ disabled: `Not available yet. ${reason}` });

/**
 * Visio's File > Options categories, in Visio's order, built from the shared Office categories
 * (`unavailableCategory`, `userInterfaceSection`, `personalizeSection`). User name and Initials
 * (the local profile, also the identity in a shared session) and Office Theme (the status bar's
 * look) are honoured; every other setting is shown disabled with the capability it is waiting for.
 */
export const VISIO_OPTION_CATEGORIES: readonly OfficeOptionCategory[] = [
	{
		id: 'general',
		label: 'General',
		description: 'General options for working with Visio.',
		sections: [
			userInterfaceSection(visio),
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
	unavailableCategory(visio, 'proofing', {
		reason: 'Not available yet. Spelling and AutoCorrect need rich text editing.',
	}),
	unavailableCategory(visio, 'save', {
		description: 'Customize how drawings are saved.',
		reason: 'Not available yet. Drawings are saved as downloaded copies; AutoRecover is not kept.',
	}),
	unavailableCategory(visio, 'language', {
		reason: 'Not available yet. The viewer is English only.',
	}),
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
						max: 99,
						...disabled('The history depth is fixed.'),
					},
					{
						kind: 'toggle',
						key: 'autoSizeText',
						label: 'Auto-size text while editing',
						...disabled('Needs core text layout.'),
					},
				],
			},
		],
	},
	unavailableCategory(visio, 'ribbon'),
	unavailableCategory(visio, 'qat', {
		reason: 'Not available yet. The Quick Access Toolbar shows Undo and Redo.',
	}),
	unavailableCategory(visio, 'addins'),
	unavailableCategory(visio, 'trust'),
];
