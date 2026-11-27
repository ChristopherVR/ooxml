/** Row numbers of the three statements, shared so each can read the others. */
import { SHEETS } from './style';

/** Sheet names as formulas spell them. */
export const INCOME = `'${SHEETS.income}'`;
export const BALANCE = `'${SHEETS.balance}'`;
export const CASH = `'${SHEETS.cash}'`;
export const COSTS = SHEETS.costs;
export const ASSUME = SHEETS.assumptions;

export const I = {
	retail: 9,
	wholesale: 10,
	online: 11,
	revenue: 12,
	cogs: 13,
	gross: 14,
	grossMargin: 15,
	labour: 17,
	rent: 18,
	utilities: 19,
	marketing: 20,
	headOffice: 21,
	operating: 22,
	ebitda: 23,
	ebitdaMargin: 24,
	depreciation: 26,
	ebit: 27,
	interest: 28,
	pretax: 29,
	tax: 30,
	net: 31,
	netMargin: 32,
} as const;

export const B = {
	cash: 9,
	receivables: 10,
	inventory: 11,
	current: 12,
	fixed: 13,
	assets: 14,
	payables: 17,
	debt: 18,
	liabilities: 19,
	capital: 22,
	retained: 23,
	equity: 24,
	funding: 25,
	check: 27,
	netDebt: 30,
	leverage: 31,
} as const;

export const F = {
	net: 9,
	depreciation: 10,
	receivables: 11,
	inventory: 12,
	payables: 13,
	operations: 14,
	fitOut: 17,
	maintenance: 18,
	extension: 19,
	investing: 20,
	borrowing: 23,
	repayments: 24,
	dividends: 25,
	financing: 26,
	change: 28,
	opening: 29,
	closing: 30,
	check: 33,
	headroom: 34,
} as const;

/** A SUM of the rows `from`..`to` of a column, as a formula. */
export const sum = (from: number, to: number) => (col: string) => `=SUM(${col}${from}:${col}${to})`;
/** The same formula for the base year and the forecast years. */
export const all = (cell: (col: string) => string) => ({ base: cell('D'), cell });
