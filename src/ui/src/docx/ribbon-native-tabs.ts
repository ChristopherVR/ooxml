import { group, stack, tool } from './ribbon-parts';
import { designControl } from './design-ribbon-controls';
import type { RibbonIcon } from './ribbon-icons';
import { buildMailings } from './mailings-ribbon';

const unavailable = (label: string, icon: RibbonIcon, inline = false) =>
	tool(label, icon, null, inline ? { inline: true } : { large: true });

/** Office's tabs and groups, with unsupported operations explicitly disabled. */
export function buildNativeTabs(panels: Map<string, HTMLElement>): void {
	const design = panels.get('Design')!;
	design.append(
		group(
			'Document Formatting',
			designControl('Themes', 'pageColor', 'theme'),
			designControl('Colors', 'fontColor', 'colors'),
			designControl('Fonts', 'changeCase', 'fonts'),
			designControl('Paragraph spacing', 'lineSpacing', 'spacing'),
			designControl('Effects', 'formatPainter', 'effects'),
			designControl('Set as Default', 'accept', 'default'),
		),
	);
	const background = panels
		.get('Layout')!
		.querySelector<HTMLElement>('[data-label="Page background"]');
	if (background) design.append(background);
	buildMailings(panels.get('Mailings')!);
	const help = tool('Help', 'help', { type: 'help' }, { large: true });
	panels.get('Help')!.append(group('Help', help));
}
