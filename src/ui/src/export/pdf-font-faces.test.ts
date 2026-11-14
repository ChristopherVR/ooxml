import { describe, expect, it, vi } from 'vitest';
import { collectLocalFontFiles, parseLocalFontFaces } from './pdf-font-faces';

const css = `
@font-face {
	font-family: "Deck Sans";
	src: url("data:font/ttf;base64,AAEAAA==") format("truetype");
	font-weight: 700;
	font-style: italic;
}
@font-face { font-family: 'Remote'; src: url("https://fonts.example/r.woff2") format("woff2"); }
@font-face { font-family: Local; src: local("Arial"), url(blob:http://x/1); }
`;

describe('local @font-face fonts for PDF embedding', () => {
	it('keeps only data: and blob: sources, with weight and style', () => {
		expect(parseLocalFontFaces(css)).toEqual([
			{ family: 'Deck Sans', bold: true, italic: true, url: 'data:font/ttf;base64,AAEAAA==' },
			{ family: 'Local', bold: false, italic: false, url: 'blob:http://x/1' },
		]);
	});

	it('decodes data: fonts without any network request', async () => {
		const fetchSpy = vi.spyOn(globalThis, 'fetch');
		try {
			const doc = {
				querySelectorAll: () => [{ textContent: css.split('@font-face { font-family: Local')[0]! }],
			};
			const files = await collectLocalFontFiles(doc);
			expect(files).toHaveLength(1);
			expect(files[0]).toMatchObject({ family: 'Deck Sans', bold: true, italic: true });
			expect([...files[0]!.data]).toEqual([0, 1, 0, 0]);
			expect(fetchSpy).not.toHaveBeenCalled();
		} finally {
			fetchSpy.mockRestore();
		}
	});
});
