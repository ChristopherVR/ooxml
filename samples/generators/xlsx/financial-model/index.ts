/** `financial-model.xlsx`: the five-year plan of Brindlemoor Bakehouse. */
import type { SampleDefinition } from '../../lib/types';
import { calculate, finish } from '../lib/finish';
import { buildAssumptions } from './assumptions';
import { buildCosts } from './costs';
import { buildCover } from './cover';
import { buildDashboard } from './dashboard';
import { decorateDashboard } from './dashboard-charts';
import { buildRevenue } from './revenue';
import { buildStatements } from './statements';
import { createBook } from './style';

async function generate(): Promise<Uint8Array> {
	const book = createBook();
	buildCover(book);
	buildAssumptions(book);
	buildRevenue(book);
	buildCosts(book);
	buildStatements(book);
	buildDashboard(book);
	const session = calculate(book);
	decorateDashboard(book, session);
	return finish(book, session, {
		title: 'Brindlemoor Bakehouse five-year financial model',
		subject: 'Financial plan FY2026 to FY2030',
		description:
			'A driver-based three-statement model with scenarios. Brindlemoor Bakehouse is fictitious.',
		keywords: 'financial model, forecast, income statement, balance sheet, cash flow, scenarios',
		company: 'Brindlemoor Bakehouse',
	});
}

export const financialModel: SampleDefinition = {
	id: 'financial-model',
	product: 'xlsx',
	fileName: 'financial-model.xlsx',
	title: 'Five-year financial model',
	description:
		'A driver-based plan for Brindlemoor Bakehouse: validated assumptions with three scenarios, a revenue and cost build, and an income statement, balance sheet and cash flow that tie, summarised on a dashboard.',
	features: [
		'Eight linked sheets with cross-sheet formulas and cached results',
		'Workbook and sheet scoped named ranges',
		'Data validation: lists, whole-number and decimal ranges, input messages, error alerts',
		'Conditional formatting: data bars, icon sets, colour scales and formula rules',
		'Custom number formats: thousands, parentheses for negatives, multiples, year labels',
		'Outlined rows and columns with summary rows, two levels deep',
		'Frozen panes, restyled cell styles and a custom theme',
		'Native charts: stacked columns with a line on a secondary axis, and a line chart',
		'Sparklines, Scenario Manager scenarios, internal and external hyperlinks',
	],
	generate,
};
