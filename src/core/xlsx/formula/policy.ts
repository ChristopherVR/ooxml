// Which formulas an automated writer (an AI assistant, an MCP client) may put in a workbook. A
// formula is read with the formula tokenizer and parser, never matched as text, so a function
// is found wherever its name appears: called, passed bare to MAP, LET or LAMBDA, or written with
// the storage prefixes Excel uses (`_xlfn._xlws.WEBSERVICE`). Text inside string literals is data,
// not a reference, and never matches. A call to an XLL add-in (`_xll.Name`) runs code outside the
// workbook, so every one is restricted.
import { parseCellInput } from '../numfmt/input';
import { normalizeFunctionName } from './ast';
import { parseTokens } from './parser';
import { tokenize, type Token } from './tokenizer';

/** Why a formula is restricted. */
export type FormulaPolicyConcern =
	/** Fetches data, opens links or runs code outside the workbook when calculated or clicked. */
	| 'external-function'
	/** Sends cell data to an AI or translation service. */
	| 'ai-function'
	/** Reads another workbook or a file path. */
	| 'external-reference'
	/** Not a formula the parser reads (DDE links, for example), so it cannot be checked. */
	| 'unreadable';

export interface FormulaPolicyIssue {
	concern: FormulaPolicyConcern;
	/** The function or reference that caused it, as written. */
	name?: string;
	/** A sentence fragment: "it uses WEBSERVICE, which can reach outside the workbook". */
	reason: string;
}

/**
 * The restricted functions, by upper-case name. Spreadsheet applications other than Excel read
 * the same files, so their network and AI functions are listed too. The Excel 4.0 macro functions
 * that a defined name can call are listed where they evaluate text as a formula (`EVALUATE`), run
 * code or programs, use DDE, or read and write files.
 */
export const RESTRICTED_FORMULA_FUNCTIONS: ReadonlyMap<
	string,
	'external-function' | 'ai-function'
> = new Map([
	...[
		'WEBSERVICE',
		'FILTERXML',
		'HYPERLINK',
		'IMAGE',
		'RTD',
		'CALL',
		'REGISTER',
		'REGISTER.ID',
		'SQL.REQUEST',
		'STOCKHISTORY',
		'PY',
		'CUBEKPIMEMBER',
		'CUBEMEMBER',
		'CUBEMEMBERPROPERTY',
		'CUBERANKEDMEMBER',
		'CUBESET',
		'CUBESETCOUNT',
		'CUBEVALUE',
		'IMPORTDATA',
		'IMPORTFEED',
		'IMPORTHTML',
		'IMPORTRANGE',
		'IMPORTXML',
		'GOOGLEFINANCE',
		// Excel 4.0 macro functions.
		'EVALUATE',
		'EXEC',
		'EXECUTE',
		'RUN',
		'INITIATE',
		'REQUEST',
		'POKE',
		'FOPEN',
		'FREAD',
		'FREADLN',
		'FWRITE',
		'FWRITELN',
		'FILES',
	].map((name) => [name, 'external-function'] as const),
	...['COPILOT', 'AI', 'TRANSLATE', 'DETECTLANGUAGE', 'GOOGLETRANSLATE'].map(
		(name) => [name, 'ai-function'] as const,
	),
]);

/** How a file stores a call to a function of an XLL add-in, upper-cased. */
const ADD_IN_PREFIX = '_XLL.';

/** Characters a sheet name cannot hold; in a reference prefix they mean a path or a workbook. */
const NOT_IN_SHEET_NAMES = /[\\/[\]:?*]/u;

function referenceIssue(token: Token): FormulaPolicyIssue | null {
	const prefix = token.prefix;
	const external =
		(prefix &&
			(prefix.book !== undefined ||
				NOT_IN_SHEET_NAMES.test(prefix.sheet) ||
				NOT_IN_SHEET_NAMES.test(prefix.sheet2 ?? ''))) ||
		// A function qualified by a sheet or workbook (`PERSONAL.XLSB!Macro(`) is defined elsewhere.
		(token.kind === 'func' && token.text.length > String(token.value).length);
	return external
		? {
				concern: 'external-reference',
				name: token.text,
				reason: 'it refers to another workbook or a file',
			}
		: null;
}

/**
 * Why `formula` (with or without its leading `=`) is restricted, or null when it only uses
 * ordinary calculation. The first issue found is returned.
 */
export function formulaPolicyIssue(formula: string): FormulaPolicyIssue | null {
	let tokens: Token[];
	try {
		tokens = tokenize(formula);
		parseTokens(tokens);
	} catch {
		return {
			concern: 'unreadable',
			reason: 'it could not be read as a formula, so it could not be checked',
		};
	}
	for (const token of tokens) {
		const issue = referenceIssue(token);
		if (issue) return issue;
		if (token.kind !== 'func' && token.kind !== 'name') continue;
		const name = normalizeFunctionName(String(token.value));
		if (name.startsWith(ADD_IN_PREFIX))
			return {
				concern: 'external-function',
				name: token.text,
				reason: `it calls the add-in function ${token.text}, which runs code outside the workbook`,
			};
		const concern = RESTRICTED_FORMULA_FUNCTIONS.get(name);
		if (concern)
			return {
				concern,
				name,
				reason:
					concern === 'ai-function'
						? `it uses ${name}, which sends data to an AI service`
						: `it uses ${name}, which can reach outside the workbook`,
			};
	}
	return null;
}

/**
 * {@link formulaPolicyIssue} for text typed into a cell: only input that the grid stores as a
 * formula (it starts with `=`) is checked, so `+WEBSERVICE(...)` and `'=x` stay plain text.
 */
export function cellInputPolicyIssue(input: string): FormulaPolicyIssue | null {
	const { formula } = parseCellInput(input);
	return formula === undefined ? null : formulaPolicyIssue(formula);
}
