// Review, View and Help tabs.
import type { RibbonTab } from '../parts.js';

export function reviewTab(): RibbonTab {
	return {
		id: 'review',
		label: 'Review',
		groups: [
			{
				id: 'proofing',
				label: 'Proofing',
				controls: [{ kind: 'button', command: 'review.statistics', size: 'large' }],
			},
			{
				id: 'comments',
				label: 'Comments',
				controls: [
					{ kind: 'button', command: 'review.new-comment', size: 'large' },
					{ kind: 'button', command: 'review.delete-comment', size: 'large' },
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'review.previous-comment', showLabel: true },
							{ kind: 'button', command: 'review.next-comment', showLabel: true },
						],
					},
					{ kind: 'button', command: 'review.show-comments', size: 'large' },
				],
			},
			{
				id: 'protect',
				label: 'Protect',
				controls: [
					{ kind: 'toggle', command: 'review.protect-sheet', size: 'large' },
					{ kind: 'toggle', command: 'review.protect-workbook', size: 'large' },
				],
			},
		],
	};
}

export function viewTab(): RibbonTab {
	return {
		id: 'view',
		label: 'View',
		groups: [
			{
				id: 'workbook-views',
				label: 'Workbook Views',
				controls: [{ kind: 'toggle', command: 'view.normal', size: 'large' }],
			},
			{
				id: 'show',
				label: 'Show',
				controls: [
					{
						kind: 'stack',
						controls: [
							{ kind: 'toggle', command: 'view.formula-bar', showLabel: true },
							{ kind: 'toggle', command: 'page.gridlines-view', showLabel: true },
							{ kind: 'toggle', command: 'page.headings-view', showLabel: true },
						],
					},
				],
			},
			{
				id: 'zoom',
				label: 'Zoom',
				controls: [
					{ kind: 'button', command: 'view.zoom', size: 'large' },
					{ kind: 'button', command: 'view.zoom-100', size: 'large' },
					{ kind: 'button', command: 'view.zoom-selection', size: 'large' },
				],
			},
			{
				id: 'window',
				label: 'Window',
				controls: [
					{
						kind: 'menu',
						label: 'Freeze Panes',
						icon: 'xl-freeze',
						size: 'large',
						items: [
							{ command: 'view.freeze-panes' },
							{ command: 'view.freeze-top-row' },
							{ command: 'view.freeze-first-column' },
							{ command: 'view.unfreeze' },
						],
					},
				],
			},
		],
	};
}

export function helpTab(): RibbonTab {
	return {
		id: 'help',
		label: 'Help',
		groups: [
			{
				id: 'help',
				label: 'Help',
				controls: [
					{ kind: 'button', command: 'help.shortcuts', size: 'large' },
					{ kind: 'button', command: 'help.feature-status', size: 'large' },
				],
			},
		],
	};
}
