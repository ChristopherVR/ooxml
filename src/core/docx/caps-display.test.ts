import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { allCapsText, smallCapsText } from './caps-display';
import { createFakeMeasurer, layoutDocumentModel } from './layout/index';
import { loadDocx } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('caps display (issue #367)', () => {
	it('drops the accents of letters All caps capitalises in French (France) only', () => {
		expect(allCapsText('Procédure élément', 'fr-FR')).toBe('PROCEDURE ELEMENT');
		expect(allCapsText('Procédure élément', 'fr-fr')).toBe('PROCEDURE ELEMENT');
		expect(allCapsText('Procédure élément', 'fr-CA')).toBe('PROCÉDURE ÉLÉMENT');
		expect(allCapsText('Procédure élément', 'en-US')).toBe('PROCÉDURE ÉLÉMENT');
		expect(allCapsText('Procédure élément', undefined)).toBe('PROCÉDURE ÉLÉMENT');
		// Capitals typed as such keep their accents; so does a decomposed one.
		expect(allCapsText('ÉLÉMENT é', 'fr-FR')).toBe('ÉLÉMENT E');
		expect(allCapsText('Éé', 'fr-FR')).toBe('ÉE');
	});

	it('drops the accents of the lowercase letters Small caps draws as capitals in fr-FR', () => {
		expect(smallCapsText('Procédure Élément', 'fr-FR')).toBe('Procedure Élement');
		expect(smallCapsText('Procédure Élément', 'fr-CA')).toBe('Procédure Élément');
		expect(smallCapsText('Procédure', 'en-US')).toBe('Procédure');
	});

	it('lays out a French (France) caps run without accents, as Word 16 prints it', async () => {
		const run = (props: string, text: string) =>
			`<w:p><w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body>${[
				run('<w:caps/><w:lang w:val="fr-FR"/>', 'Procédure élément'),
				run('<w:caps/><w:lang w:val="fr-CA"/>', 'Procédure élément'),
				run('<w:smallCaps/><w:lang w:val="fr-FR"/>', 'Procédure élément'),
				run('<w:lang w:val="fr-FR"/>', 'PROCÉDURE ÉLÉMENT'),
			].join(
				'',
			)}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
		);
		const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const { pages } = layoutDocumentModel(model, createFakeMeasurer());
		const lines = pages[0]!.columns[0]!.blocks.map((block) =>
			block.kind === 'paragraph'
				? block.lines.flatMap((line) => line.fragments.map((fragment) => fragment.text)).join('')
				: '',
		);
		expect(lines).toEqual([
			'PROCEDURE ELEMENT',
			'PROCÉDURE ÉLÉMENT',
			'Procedure element',
			'PROCÉDURE ÉLÉMENT',
		]);
	});
});
