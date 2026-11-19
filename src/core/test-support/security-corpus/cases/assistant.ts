import { type AiProviderSpec, normalizeAiPolicy, userConnectionRefusal } from '../../../ai/policy';
import { createXlsxFromSheets } from '../../../automation/outline';
import { xlsxTools } from '../../../automation/tools/index';
import { createXlsx, readXlsxFilledCells, setXlsxCells } from '../../../automation/xlsx';
import { loadXlsx } from '../../../xlsx/read/load';
import { saveXlsx } from '../../../xlsx/write/save';
import type { SecurityCase } from '../case';

/**
 * Formulas that reach outside the workbook when Excel calculates them, in the forms a text filter
 * missed: a function passed as a value, a namespaced storage name, and another workbook's cells.
 */
const OUTWARD_FORMULAS = [
	'=LET(f,WEBSERVICE,f(A1))',
	'=MAP(A1:A2,WEBSERVICE)',
	'=_xlfn._xlws.WEBSERVICE(A1)',
	"='C:\\Data\\[Other.xlsx]Sheet1'!A1",
];

/** The automation write paths an assistant or an MCP client reaches. */
const WRITE_PATHS: Record<string, (input: string) => Promise<unknown>> = {
	'set cells': async (input) =>
		setXlsxCells(await createXlsx(['S']), 0, [{ address: 'B1', input }]),
	'new workbook': (input) => createXlsxFromSheets([{ name: 'S', rows: [['a', input]] }]),
	'set cells tool': async (input) => {
		const tool = xlsxTools.find((candidate) => candidate.name === 'xlsx_set_cells');
		if (tool?.kind !== 'edit') throw new Error('no set cells tool');
		return tool.run(await createXlsx(['S']), { sheetIndex: 0, cells: [{ address: 'B1', input }] });
	},
};

const PROVIDERS: AiProviderSpec[] = [
	{ id: 'anthropic', name: 'Anthropic', protocol: 'anthropic', endpoint: 'https://a.test/v1' },
	{ id: 'openrouter', name: 'OpenRouter', protocol: 'openai', endpoint: 'https://o.test/v1' },
];

export const cases: readonly SecurityCase[] = [
	{
		name: 'formulas that reach outside the workbook are refused on every automation write path',
		limits: { ms: 5_000, heapMb: 128 },
		async run() {
			const written: string[] = [];
			for (const [path, write] of Object.entries(WRITE_PATHS))
				for (const formula of OUTWARD_FORMULAS)
					await write(formula).then(
						() => written.push(`${path}: ${formula}`),
						() => undefined,
					);
			return written;
		},
		expected: [],
	},
	{
		name: 'hidden sheets, rows and columns stay out of what the assistant reads by default',
		limits: { ms: 5_000, heapMb: 128 },
		async run() {
			const book = await loadXlsx(
				await createXlsxFromSheets([
					{
						name: 'Public',
						rows: [
							['shown', 'shown', 'canary-column'],
							['canary-row', 1, 2],
							['shown', 3, 4],
						],
					},
					{ name: 'VeryHidden', rows: [['canary-very-hidden']] },
					{ name: 'Hidden', rows: [['canary-hidden']] },
				]),
			);
			const [visible, veryHidden, hidden] = book.sheets;
			veryHidden!.state = 'veryHidden';
			hidden!.state = 'hidden';
			visible!.rowInfo.set(1, { hidden: true });
			visible!.columns.push({ min: 2, max: 2, hidden: true });
			const policy = normalizeAiPolicy({}, PROVIDERS);
			const read = await readXlsxFilledCells(await saveXlsx(book), {
				hidden: policy.context.includeHidden ? 'include' : 'omit',
			});
			return JSON.stringify(read).match(/canary[\w-]*/gu) ?? [];
		},
		expected: [],
	},
	{
		name: 'a saved connection to a provider the workspace does not allow is refused',
		limits: { ms: 2_000, heapMb: 64 },
		run() {
			const policy = normalizeAiPolicy({ userProviders: ['anthropic'] }, PROVIDERS);
			const refusal = (provider: string, endpoint: string) =>
				userConnectionRefusal(policy, { provider, endpoint }, PROVIDERS) !== null;
			return {
				otherProvider: refusal('openrouter', 'https://o.test/v1'),
				plainHttp: refusal('anthropic', 'http://a.test/v1'),
				credentials: refusal('anthropic', 'https://user:secret@a.test/v1'),
			};
		},
		expected: { otherProvider: true, plainHttp: true, credentials: true },
	},
];
