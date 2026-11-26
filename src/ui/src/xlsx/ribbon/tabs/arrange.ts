import type { RibbonGroup } from 'ooxml-core/xlsx/ui';
import { ALIGNMENTS } from '../../commands/arrange';

export function arrangeGroup(): RibbonGroup {
	return {
		id: 'arrange',
		label: 'Arrange',
		controls: [
			{ kind: 'button', command: 'arrange.forward', size: 'large' },
			{ kind: 'button', command: 'arrange.backward', size: 'large' },
			{ kind: 'button', command: 'arrange.selection-pane', size: 'large' },
			{
				kind: 'stack',
				controls: [
					{
						kind: 'menu',
						label: 'Align',
						items: ALIGNMENTS.map(([id]) => ({ command: 'arrange.align-' + id })),
					},
					{
						kind: 'menu',
						label: 'Group',
						items: [{ command: 'arrange.group' }, { command: 'arrange.ungroup' }],
					},
					{
						kind: 'menu',
						label: 'Rotate',
						items: [
							{ command: 'arrange.rotate-right' },
							{ command: 'arrange.rotate-left' },
							{ command: 'arrange.flip-horizontal' },
							{ command: 'arrange.flip-vertical' },
						],
					},
				],
			},
		],
	};
}
