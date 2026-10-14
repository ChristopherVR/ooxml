import { group, stack, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';

const unavailable = (label: string, icon: RibbonIcon, inline = false) =>
	tool(label, icon, null, inline ? { inline: true } : { large: true });

/** Office's tabs and groups, with unsupported operations explicitly disabled. */
export function buildNativeTabs(panels: Map<string, HTMLElement>): void {
	const design = panels.get('Design')!;
	design.append(
		group(
			'Document Formatting',
			unavailable('Themes', 'pageColor'),
			unavailable('Colors', 'fontColor'),
			unavailable('Fonts', 'changeCase'),
			unavailable('Paragraph spacing', 'lineSpacing'),
			unavailable('Effects', 'formatPainter'),
			unavailable('Set as Default', 'accept'),
		),
	);
	const background = panels
		.get('Layout')!
		.querySelector<HTMLElement>('[data-label="Page background"]');
	if (background) design.append(background);
	panels
		.get('Mailings')!
		.append(
			group('Create', unavailable('Envelopes', 'envelope'), unavailable('Labels', 'table')),
			group(
				'Start Mail Merge',
				unavailable('Start Mail Merge', 'envelope'),
				unavailable('Select Recipients', 'select'),
				unavailable('Edit Recipient List', 'table'),
			),
			group(
				'Write & Insert Fields',
				unavailable('Highlight Merge Fields', 'highlight'),
				unavailable('Address Block', 'textBox'),
				unavailable('Greeting Line', 'textBox'),
				unavailable('Insert Merge Field', 'field'),
				stack(
					unavailable('Rules', 'field', true),
					unavailable('Match Fields', 'field', true),
					unavailable('Update Labels', 'updateTable', true),
				),
			),
			group(
				'Preview Results',
				unavailable('Preview Results', 'find'),
				stack(
					unavailable('Find Recipient', 'find', true),
					unavailable('Check for Errors', 'spelling', true),
				),
			),
			group('Finish', unavailable('Finish & Merge', 'envelope')),
		);
	const help = tool('Help', 'help', { type: 'help' }, { large: true });
	panels.get('Help')!.append(group('Help', help));
}
