import { icon, type RibbonTab } from 'ooxml-core/xlsx/ui';
import { hasSparklines } from '../../commands/sparklines';

export function sparklineTab(): RibbonTab {
	return {
		id: 'sparkline',
		label: 'Sparkline',
		contextual: hasSparklines,
		deferControls: true,
		groups: [
			{
				id: 'sparkline',
				label: 'Sparkline',
				controls: [{ kind: 'button', command: 'sparkline.settings', size: 'large' }],
			},
			{
				id: 'group',
				label: 'Group',
				controls: [
					{ kind: 'button', command: 'sparkline.group', size: 'large' },
					{ kind: 'button', command: 'sparkline.ungroup', size: 'large' },
					{
						kind: 'menu',
						label: 'Clear Sparklines',
						icon: icon('clear'),
						size: 'large',
						items: [{ command: 'sparkline.clear-selected' }, { command: 'sparkline.clear-groups' }],
					},
				],
			},
		],
	};
}
