// Draw tab (after Insert), grouped as Microsoft 365's: Drawing Tools, Stencils, Convert, Replay.
import type { RibbonTab } from 'ooxml-core/xlsx/ui';
import { renderPenStrip } from '../../commands/draw';

export function drawTab(): RibbonTab {
	return {
		id: 'draw',
		label: 'Draw',
		groups: [
			{
				id: 'drawing-tools',
				label: 'Drawing Tools',
				controls: [
					{ kind: 'toggle', command: 'draw.select', size: 'large' },
					{ kind: 'toggle', command: 'draw.lasso', size: 'large' },
					{ kind: 'toggle', command: 'draw.eraser', size: 'large' },
					{ kind: 'element', command: 'draw.pens', render: renderPenStrip },
					{
						kind: 'menu',
						label: 'Add Pen',
						icon: 'xl-addPen',
						size: 'large',
						items: [
							{ command: 'draw.add-pen', arg: 'pen', label: 'Pen' },
							{ command: 'draw.add-pen', arg: 'pencil', label: 'Pencil' },
							{ command: 'draw.add-pen', arg: 'highlighter', label: 'Highlighter' },
						],
					},
					{ kind: 'toggle', command: 'draw.touch', size: 'large' },
				],
			},
			{
				id: 'stencils',
				label: 'Stencils',
				controls: [{ kind: 'button', command: 'draw.ruler', size: 'large' }],
			},
			{
				id: 'convert',
				label: 'Convert',
				controls: [
					{ kind: 'button', command: 'draw.ink-to-shape', size: 'large' },
					{ kind: 'button', command: 'draw.ink-to-math', size: 'large' },
				],
			},
			{
				id: 'replay',
				label: 'Replay',
				controls: [{ kind: 'toggle', command: 'draw.replay', size: 'large' }],
			},
		],
	};
}
