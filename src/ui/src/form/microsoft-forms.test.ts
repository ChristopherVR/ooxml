// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { microsoftFormsResponseUrl, pickMicrosoftForm } from './microsoft-forms';

afterEach(() => document.body.replaceChildren());

it.each([
	'https://forms.office.com/r/Demo123',
	'https://forms.cloud.microsoft/pages/responsepage.aspx?id=abc123',
])('accepts Microsoft Forms response URLs (%s)', (url) => {
	expect(microsoftFormsResponseUrl(url)).toBe(url);
});

it.each([
	'javascript:alert(1)',
	'https://forms.office.com.example.test/r/abc',
	'https://forms.office.com@elsewhere.example/r/abc',
	'https://forms.office.com/Pages/DesignPageV2.aspx',
	'https://forms.office.com/pages/responsepage.aspx',
	'https://forms.office.com/pages/responsepage.aspx?id=',
	'http://forms.office.com/r/abc',
])('rejects other destinations (%s)', (url) => {
	expect(microsoftFormsResponseUrl(url)).toBeUndefined();
});

it('returns the response link only after the form is submitted', async () => {
	const result = pickMicrosoftForm(document.body, (label) => label);
	const form = document.body.querySelector('form')!;
	form.querySelector<HTMLInputElement>('input[type="url"]')!.value =
		'https://forms.office.com/r/Demo123';
	form.requestSubmit();
	await expect(result).resolves.toEqual({
		title: 'Open form',
		url: 'https://forms.office.com/r/Demo123',
	});
});
