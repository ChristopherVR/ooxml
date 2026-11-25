import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx } from './index';
import type { ThemeCatalog } from './theme-model';
import { must } from './test-support/access';

it('creates a theme part, then patches the related source with its extensions intact', async () => {
	const theme: ThemeCatalog = {
		colors: { accent1: '123456' },
		fonts: { major: { latin: 'Georgia' }, minor: { latin: 'Calibri' } },
		colorMapping: { tx1: 'dk2' },
	};
	const zip = await JSZip.loadAsync(await saveDocx({ ...createDocument(), theme }));
	const path = 'word/theme/custom.xml';
	const settingsPath = 'word/theme/settings.xml';
	const xml = must(await zip.file('word/theme/theme1.xml')?.async('string'));
	zip
		.remove('word/theme/theme1.xml')
		.file(
			path,
			xml
				.replace('</a:theme>', '<a:extLst/></a:theme>')
				.replace('</a:fmtScheme>', '<a:extLst/></a:fmtScheme>'),
		);
	const settings = must(await zip.file('word/settings.xml')?.async('string'));
	zip.remove('word/settings.xml').file(settingsPath, settings);
	for (const file of ['[Content_Types].xml', 'word/_rels/document.xml.rels']) {
		const xml = must(await zip.file(file)?.async('string'));
		zip.file(
			file,
			xml
				.replace('theme/theme1.xml', 'theme/custom.xml')
				.replace('settings.xml', 'theme/settings.xml'),
		);
	}
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const original = must(loaded.model.theme);
	expect(original).toMatchObject(theme);
	const { sourceXml: _source, ...matrix } = must(original.formatScheme);
	const updated: ThemeCatalog = {
		...original,
		colors: { ...original.colors, accent1: 'AABBCC' },
		fonts: { ...original.fonts, minor: { latin: 'Verdana' } },
		colorMapping: { bg1: 'lt2' },
		formatScheme: { ...matrix, name: 'Document Format' },
	};
	const bytes = await loaded.save({ ...loaded.model, theme: updated });
	const output = await JSZip.loadAsync(bytes);
	expect(output.file('word/theme/theme1.xml')).toBeNull();
	expect(output.file('word/settings.xml')).toBeNull();
	expect(await output.file(settingsPath)?.async('string')).toContain('w:bg1="light2"');
	expect((await output.file(path)?.async('string'))?.match(/<a:extLst\/>/g)).toHaveLength(2);
	expect((await loadDocx(bytes)).model.theme).toMatchObject(updated);
});
