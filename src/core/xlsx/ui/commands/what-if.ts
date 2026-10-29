// Data tab, Forecast group: What-If Analysis (Scenario Manager, Goal Seek, Data Table) and the
// Forecast Sheet command (a dialog that builds the forecast on a new sheet).
import type { Command } from '../commands';
import { icon } from './icons';
import { editing } from './util';

export function whatIfCommands(): Command[] {
	return [
		editing({
			id: 'data.scenario-manager',
			label: 'Scenario Manager...',
			icon: icon('scenarios'),
			run: (ctx) => void ctx.dialogs.open('scenario-manager'),
		}),
		editing({
			id: 'data.goal-seek',
			label: 'Goal Seek...',
			icon: icon('goalSeek'),
			run: (ctx) => void ctx.dialogs.open('goal-seek'),
		}),
		editing({
			id: 'data.data-table',
			label: 'Data Table...',
			icon: icon('dataTable'),
			run: (ctx) => void ctx.dialogs.open('data-table'),
		}),
		editing({
			id: 'data.forecast-sheet',
			label: 'Forecast Sheet',
			icon: icon('forecast'),
			run: (ctx) => void ctx.dialogs.open('forecast-sheet'),
		}),
	];
}
