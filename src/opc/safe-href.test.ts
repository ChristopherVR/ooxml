import { describe, expect, it } from 'vitest';

import {
	hyperlinkPolicy,
	isOpenableHyperlinkHref,
	isSafeHyperlinkHref,
	isStorableHyperlinkHref,
} from './safe-href.js';

// Built without the literal token so script-url lint rules stay quiet.
const JS = `${'java'}script`;
const ch = (code: number): string => String.fromCharCode(code);

describe('hyperlinkPolicy: openable', () => {
	it.each([
		'https://example.com',
		'http://example.com',
		'HTTPS://EXAMPLE.COM',
		'mailto:user@example.com',
		'  https://example.com  ',
		`https://example.com/${JS}/docs`,
		'https://example.com/a b',
	])('opens and stores %s', (href) => {
		expect(hyperlinkPolicy(href)).toEqual({ store: true, open: true });
		expect(isOpenableHyperlinkHref(href)).toBe(true);
		expect(isStorableHyperlinkHref(href)).toBe(true);
	});
});

describe('hyperlinkPolicy: store only', () => {
	it.each([
		'ftp://files.example.com',
		'file:///C:/Reports/q1.xlsx',
		'FILE://server/share/a.xlsx',
		'#Sheet2!A1',
		"#'My Sheet'!B3",
		'Sheet2!A1',
		'/page/about',
		'../reports/q1.xlsx',
		'q1.xlsx',
		'\\\\server\\share\\q1.xlsx',
		'C:\\Reports\\q1.xlsx',
		'//example.com/x',
	])('stores but does not open %s', (href) => {
		expect(hyperlinkPolicy(href)).toEqual({ store: true, open: false, reason: 'not-openable' });
	});
});

describe('hyperlinkPolicy: rejected', () => {
	const scripts = [
		`${JS}:alert(1)`,
		`${JS.toUpperCase()}:alert(1)`,
		`JaVaScRiPt:alert(1)`,
		`  ${JS}:alert(1)`,
		`${ch(1)}${JS}:alert(1)`,
		'data:text/html,<h1>XSS</h1>',
		'DATA:text/html;base64,PHNjcmlwdD4=',
		"vbscript:MsgBox('XSS')",
		'mhtml:file://C:/test.mht',
		`x-${JS}:alert(1)`,
		'livescript:alert(1)',
		'ms-its:c:/x.chm::/a.htm',
		'blob:https://example.com/uuid',
	];
	it.each(scripts)('rejects script scheme %s', (href) => {
		expect(hyperlinkPolicy(href)).toEqual({ store: false, open: false, reason: 'script-scheme' });
	});

	it.each([
		['zero-width space', `java${ch(0x200b)}script:alert(1)`],
		['zero-width non-joiner', `java${ch(0x200c)}script:alert(1)`],
		['zero-width joiner', `java${ch(0x200d)}script:alert(1)`],
		['byte order mark', `${ch(0xfeff)}${JS}:alert(1)`],
		['NUL', `java${ch(0)}script:alert(1)`],
		['tab', `java\tscript:alert(1)`],
		['newline', `java\nscript:alert(1)`],
		['carriage return', `java\rscript:alert(1)`],
		['space', 'java script:alert(1)'],
		['soft hyphen', `java${ch(0xad)}script:alert(1)`],
		['word joiner', `java${ch(0x2060)}script:alert(1)`],
		['bidi override', `${ch(0x202e)}${JS}:alert(1)`],
		['tab inside data', 'da\tta:text/html,x'],
	])('rejects a script scheme hidden by %s', (_name, href) => {
		expect(hyperlinkPolicy(href).store).toBe(false);
		expect(hyperlinkPolicy(href).open).toBe(false);
		expect(hyperlinkPolicy(href).reason).toBe('script-scheme');
	});

	it('rejects an allowed scheme split by ignorable characters', () => {
		for (const href of ['ht\ttp://example.com', `ht${ch(0x200b)}tps://example.com`])
			expect(hyperlinkPolicy(href)).toEqual({
				store: false,
				open: false,
				reason: 'obfuscated-scheme',
			});
	});

	it('rejects control characters anywhere in the href', () => {
		for (const href of [`https://example.com/${ch(0)}`, 'https://a\nb', `#Sheet1${ch(7)}!A1`])
			expect(hyperlinkPolicy(href).reason).toBe('control-characters');
	});

	it('rejects unknown schemes', () => {
		for (const href of ['tel:+123', 'ms-word:ofe|u|https://x', 'chrome://settings', 'about:blank'])
			expect(hyperlinkPolicy(href)).toEqual({
				store: false,
				open: false,
				reason: 'unsupported-scheme',
			});
	});

	it('rejects empty and missing values', () => {
		for (const href of ['', '   ', ch(0x200b), undefined, null])
			expect(hyperlinkPolicy(href)).toEqual({ store: false, open: false, reason: 'empty' });
		expect(isStorableHyperlinkHref(undefined)).toBe(false);
		expect(isOpenableHyperlinkHref(null)).toBe(false);
	});
});

describe('isSafeHyperlinkHref keeps its strict behaviour', () => {
	it('matches the open column for plain inputs and still rejects relative targets', () => {
		expect(isSafeHyperlinkHref('https://a')).toBe(true);
		expect(isSafeHyperlinkHref('mailto:x@y.z')).toBe(true);
		expect(isSafeHyperlinkHref('#Sheet2!A1')).toBe(false);
		expect(isSafeHyperlinkHref('ftp://a')).toBe(false);
	});
});
