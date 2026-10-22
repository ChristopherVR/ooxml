import type { RibbonControlId } from './customization';
import type { RibbonGroupView } from './ribbon-command-view';

/** Shared Help ordering and presentation. Hosts own dialog availability and actions. */
export const HELP_RIBBON_COMMANDS: readonly {
	id: RibbonControlId;
	labelKey: string;
	icon: string;
}[] = [
	{ id: 'help.help.options', labelKey: 'pptx.settings.title', icon: 'settings' },
	{
		id: 'help.help.keyboardShortcuts',
		labelKey: 'pptx.settings.keyboardShortcuts',
		icon: 'keyboard',
	},
	{
		id: 'help.help.accessibility',
		labelKey: 'pptx.ribbon.accessibilityCheck',
		icon: 'accessibility',
	},
];

/** Which Help commands the binding can open; the others are left out of the tab. */
export interface HelpRibbonState {
	/** The Options dialog is available (`isDialogAvailable(customization, 'options')`). */
	options: boolean;
	/**
	 * The keyboard shortcuts reference can open. Default true; a binding whose shortcuts are a tab
	 * of the Options dialog passes the Options availability.
	 */
	shortcuts?: boolean;
}

/** The Help tab for `pptx-ui-ribbon-section`: one group of compact commands, in shared order. */
export function buildHelpRibbon(
	t: (key: string) => string,
	state: HelpRibbonState,
): RibbonGroupView[] {
	const available: Partial<Record<RibbonControlId, boolean>> = {
		'help.help.options': state.options,
		'help.help.keyboardShortcuts': state.shortcuts !== false,
	};
	return [
		{
			id: 'help.help',
			label: t('pptx.ribbon.tab.help'),
			commands: HELP_RIBBON_COMMANDS.filter((command) => available[command.id] !== false).map(
				(command) => ({
					id: command.id,
					label: t(command.labelKey),
					icon: command.icon,
					compact: true,
				}),
			),
		},
	];
}
