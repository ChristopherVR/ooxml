import { filePopoutDetail, filePopoutUrl } from './file-popout.js';

describe('file viewer route', () => {
	it('preserves a signed URL in the fragment without adding it to request parameters', () => {
		const detail = {
			attachment: { name: 'Budget.xlsx', kind: 'xlsx' as const },
			url: 'https://files.test/Budget.xlsx?signature=a%2Bb&expires=123',
		};
		const link = filePopoutUrl(detail, 'https://app.test/demo?local=1#old')!;
		expect(new URL(link).searchParams.get('openteams-file')).toBe('1');
		expect(new URL(link).search).not.toContain('signature');
		expect(filePopoutDetail(link)).toMatchObject(detail);
	});
	it('rejects malformed or non-web preview payloads', () => {
		for (const hash of [
			'broken',
			encodeURIComponent('null'),
			encodeURIComponent(
				JSON.stringify({
					attachment: { name: 'Book.xlsx', kind: 'xlsx' },
					url: 'javascript:alert(1)',
				}),
			),
			'a'.repeat(32_769),
		])
			expect(filePopoutDetail(`https://app.test/?openteams-file=1#${hash}`)).toBeNull();
		expect(filePopoutDetail('https://app.test/#plain-workspace-fragment')).toBeNull();
		expect(
			filePopoutUrl(
				{ attachment: { name: 'Book.xlsx', kind: 'xlsx' }, url: undefined },
				'https://app.test',
			),
		).toBeNull();
	});
});
