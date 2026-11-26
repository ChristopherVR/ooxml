import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { RibbonAddInTab } from 'ooxml-core/ribbon';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import type { AddInDialogHost } from './add-in-dialog';
import type { addInLauncher } from './add-in-launch';
import { addInIcon } from './add-in-icons';
import { addInIssueLabel } from './add-in-local-dialog';
import { syncInstalledRibbonTabs } from './add-in-tabs';

export const MANIFEST_TAB_PREFIX = 'installed-add-in.';

export function addInManifestTabs(
	host: AddInDialogHost,
	installed: ReadonlyMap<string, LocalAddIn>,
	launcher: ReturnType<typeof addInLauncher>,
	t: OfficeServiceText,
) {
	const tabs: RibbonAddInTab[] = [...installed.values()].flatMap((addIn) =>
		addIn.manifest.tabs
			.filter(
				({ groups }) =>
					!groups.every(({ commands }) =>
						commands.every(
							(command) =>
								command.kind === 'inactive' &&
								command.issue.kind === 'unsupported' &&
								command.issue.feature === 'built-in ribbon tab',
						),
					),
			)
			.map((tab) => ({
				id: `${MANIFEST_TAB_PREFIX}${addIn.manifest.id}.${tab.id}`,
				label: tab.label,
				groups: tab.groups.map((group) => ({
					label: group.label,
					commands: group.commands.map((command, index) => {
						const id = `${group.id}.${index}`;
						const iconImage = addInIcon(addIn.manifest);
						const common = {
							id,
							label: command.label,
							icon: 'addIns',
							...(iconImage && { iconImage }),
						};
						if (command.kind === 'inactive') {
							const { issue } = command;
							return {
								...common,
								disabled: true,
								title: addInIssueLabel(issue, t),
							};
						}
						const { reason } = launcher.pane(addIn, command.url);
						return reason
							? { ...common, disabled: true, title: reason }
							: {
									...common,
									run: () => void launcher.open(addIn, command.url),
								};
					}),
				})),
			})),
	);
	const ribbon = host.querySelector<HTMLElement>('office-ui-ribbon, pptx-ui-ribbon-tabs');
	if (!ribbon) return;
	if ('setInstalledAddInTabs' in ribbon && typeof ribbon.setInstalledAddInTabs === 'function')
		ribbon.setInstalledAddInTabs(tabs);
	else syncInstalledRibbonTabs(ribbon, tabs);
}
