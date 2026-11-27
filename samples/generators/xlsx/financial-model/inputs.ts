/** The validation rules of the model inputs and the scenario drivers with their three cases. */
import type { DataValidation } from '../../../../src/core/xlsx/model';
import { FMT } from './style';

export const SCENARIOS = ['Base', 'Expansion', 'Cautious'] as const;

export type Rule = Omit<DataValidation, 'ranges'>;
export const whole = (min: number, max: number, title: string, what: string): Rule => ({
	type: 'whole',
	operator: 'between',
	formula1: String(min),
	formula2: String(max),
	allowBlank: false,
	showInputMessage: true,
	promptTitle: title,
	prompt: `${what} Whole number from ${min} to ${max}.`,
	showErrorMessage: true,
	errorStyle: 'stop',
	errorTitle: 'Value out of range',
	error: `Enter a whole number between ${min} and ${max}.`,
});
export const decimal = (
	min: number,
	max: number,
	title: string,
	what: string,
	errorStyle: 'stop' | 'warning' = 'stop',
): Rule => ({
	type: 'decimal',
	operator: 'between',
	formula1: String(min),
	formula2: String(max),
	allowBlank: false,
	showInputMessage: true,
	promptTitle: title,
	prompt: what,
	showErrorMessage: true,
	errorStyle,
	errorTitle: errorStyle === 'stop' ? 'Value out of range' : 'Unusual value',
	error:
		errorStyle === 'stop'
			? `Enter a value between ${min} and ${max}.`
			: `Values outside ${Math.round(min * 100)}% to ${Math.round(max * 100)}% are unusual for this driver. Keep it?`,
});
export const amount = (title: string, what: string): Rule => ({
	type: 'decimal',
	operator: 'greaterThanOrEqual',
	formula1: '0',
	allowBlank: false,
	showInputMessage: true,
	promptTitle: title,
	prompt: what,
	showErrorMessage: true,
	errorStyle: 'stop',
	errorTitle: 'Amount cannot be negative',
	error: 'Enter zero or a positive amount in whole dollars.',
});

export interface Driver {
	title: string;
	unit: string;
	fmt: string;
	/** Base, Expansion and Cautious values for the five forecast years. */
	values: [number[], number[], number[]];
	actual?: number;
	rule: Rule;
}

export const DRIVERS: Driver[] = [
	{
		title: 'New store openings',
		unit: 'stores',
		fmt: FMT.count,
		values: [
			[1, 2, 2, 2, 2],
			[2, 3, 3, 4, 4],
			[1, 1, 1, 1, 1],
		],
		rule: whole(0, 12, 'New store openings', 'Bakeries opened during the year.'),
	},
	{
		title: 'Like-for-like transaction growth',
		unit: '% a year',
		fmt: FMT.percent,
		values: [
			[0.025, 0.025, 0.02, 0.02, 0.02],
			[0.035, 0.035, 0.03, 0.03, 0.025],
			[0.01, 0.01, 0.015, 0.015, 0.015],
		],
		rule: decimal(
			-0.1,
			0.15,
			'Like-for-like growth',
			'Growth in weekly transactions per established store.',
			'warning',
		),
	},
	{
		title: 'Average ticket inflation',
		unit: '% a year',
		fmt: FMT.percent,
		values: [
			[0.03, 0.028, 0.025, 0.025, 0.025],
			[0.032, 0.03, 0.028, 0.026, 0.025],
			[0.02, 0.02, 0.02, 0.022, 0.022],
		],
		rule: decimal(
			-0.1,
			0.15,
			'Ticket inflation',
			'Price and mix change of the average basket.',
			'warning',
		),
	},
	{
		title: 'Ingredient cost',
		unit: '% of revenue',
		fmt: FMT.percent,
		values: [
			[0.27, 0.268, 0.266, 0.265, 0.264],
			[0.271, 0.268, 0.265, 0.262, 0.26],
			[0.275, 0.276, 0.276, 0.275, 0.275],
		],
		actual: 0.272,
		rule: decimal(0.1, 0.6, 'Ingredient cost', 'Flour, dairy, fruit and other ingredients.'),
	},
	{
		title: 'Store labour',
		unit: '% of retail',
		fmt: FMT.percent,
		values: [
			[0.312, 0.31, 0.308, 0.306, 0.305],
			[0.314, 0.311, 0.308, 0.305, 0.302],
			[0.316, 0.317, 0.318, 0.318, 0.318],
		],
		actual: 0.314,
		rule: decimal(
			0.1,
			0.6,
			'Store labour',
			'Bakers and counter staff, as a share of retail revenue.',
		),
	},
];
