/**
 * The three statements. They tie by construction: the cash flow is derived from the income
 * statement and the movements of the balance sheet, whose cash line reads the closing cash back.
 */
import type { Book } from '../lib/book';
import { buildBalance } from './balance';
import { buildCashFlow } from './cash-flow';
import { buildIncome } from './income';
import { B, BALANCE, CASH, F, I, INCOME } from './rows';
import { SHEETS } from './style';

export function buildStatements(book: Book): void {
	buildIncome(book);
	buildBalance(book);
	buildCashFlow(book);
	const span = (sheet: string, row: number) => `${sheet}!$E$${row}:$I$${row}`;
	book.wb.definedNames.push(
		{ name: 'Revenue_Total', formula: span(INCOME, I.revenue) },
		{ name: 'Gross_Profit', formula: span(INCOME, I.gross) },
		{ name: 'EBITDA', formula: span(INCOME, I.ebitda) },
		{ name: 'Net_Income', formula: span(INCOME, I.net) },
		{ name: 'Closing_Cash', formula: span(CASH, F.closing) },
		{ name: 'Bank_Debt', formula: span(BALANCE, B.debt) },
		// Sheet-scoped names: each statement owns its own check row.
		{
			name: 'Check',
			formula: span(BALANCE, B.check),
			localSheet: book.sheet(SHEETS.balance).index,
		},
		{ name: 'Check', formula: span(CASH, F.check), localSheet: book.sheet(SHEETS.cash).index },
	);
}
