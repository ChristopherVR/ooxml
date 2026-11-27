import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { eighthPoints, signedTwips, twips } from './units';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(borders: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${borders}<w:pgNumType w:start="1"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const xmlOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';
const WITH =
	'<w:pgBorders w:offsetFrom="text" w:display="firstPage"><w:top w:val="single" w:sz="8" w:space="24" w:color="FF0000"/><w:left w:val="single" w:sz="4" w:space="4" w:color="auto" w:art="apples"/></w:pgBorders>';

describe('page borders', () => {
	it('parses sides, spacing, art, offset and display', async () => {
		const { model } = await loadDocx(await fixture(WITH));
		expect(model.sections![0]!.pageBorders).toMatchObject({
			offsetFrom: 'text',
			display: 'firstPage',
			top: { style: 'single', sizeEighthPoints: 8, color: '#FF0000', spacePoints: 24 },
			left: { spacePoints: 4, art: 'apples' },
		});
		expect(model.sections![0]!.pageBorders!.bottom).toBeUndefined();
	});

	it('rewrites only the changed side and keeps the art side and attributes', async () => {
		const loaded = await loadDocx(await fixture(WITH));
		const borders = loaded.model.sections![0]!.pageBorders!;
		borders.top = { ...borders.top!, sizeEighthPoints: eighthPoints(18), color: '#00ff00' };
		borders.bottom = { style: 'double', sizeEighthPoints: eighthPoints(6), spacePoints: 10 };
		const xml = await xmlOf(await loaded.save());
		expect(xml).toMatch(
			/<w:pgBorders w:offsetFrom="text" w:display="firstPage"><w:top [^>]*w:sz="18"[^>]*w:color="00FF00"[^>]*\/><w:left [^>]*w:art="apples"[^>]*\/><w:bottom w:val="double" w:sz="6" w:space="10" w:color="auto"\/><\/w:pgBorders><w:pgNumType/,
		);
	});

	it('creates pgBorders in schema order and removes it again', async () => {
		const loaded = await loadDocx(await fixture(''));
		loaded.model.sections![0]!.pageBorders = {
			right: { style: 'single', sizeEighthPoints: eighthPoints(4), spacePoints: 24 },
			offsetFrom: 'page',
		};
		const saved = await loaded.save();
		const xml = await xmlOf(saved);
		expect(xml).toMatch(
			/<w:pgMar [^>]*\/><w:pgBorders w:offsetFrom="page"><w:right [^>]*\/><\/w:pgBorders><w:pgNumType/,
		);
		const reopened = await loadDocx(saved);
		delete reopened.model.sections![0]!.pageBorders;
		expect(await xmlOf(await reopened.save())).not.toContain('pgBorders');
	});

	it('keeps the package byte-exact when borders are untouched', async () => {
		const bytes = await fixture(WITH);
		expect(await (await loadDocx(bytes)).save()).toEqual(bytes);
	});

	it('keeps the borders of the following section out of a new break before it', async () => {
		const { saveDocx } = await import('./save');
		const section = (endsAtBlockId: string) => ({
			endsAtBlockId,
			type: 'nextPage' as const,
			pageWidthTwips: twips(12240),
			pageHeightTwips: twips(15840),
			orientation: 'portrait' as const,
			marginTopTwips: signedTwips(1440),
			marginRightTwips: twips(1440),
			marginBottomTwips: signedTwips(1440),
			marginLeftTwips: twips(1440),
			columns: { count: 1, equalWidth: true },
		});
		const bytes = await saveDocx({
			blocks: [
				{ type: 'paragraph', id: 'a', runs: [{ text: 'Plain' }] },
				{ type: 'paragraph', id: 'b', runs: [{ text: 'Bordered' }] },
			],
			page: {
				width: 816,
				height: 1056,
				marginTop: 96,
				marginRight: 96,
				marginBottom: 96,
				marginLeft: 96,
			},
			warnings: [],
			sections: [
				section('a'),
				{
					...section('b'),
					pageBorders: {
						top: { style: 'single', sizeEighthPoints: eighthPoints(8), spacePoints: 24 },
						offsetFrom: 'page',
					},
				},
			],
		});
		const xml = await xmlOf(bytes);
		const paragraphSection = /<w:pPr>.*?<w:sectPr>(.*?)<\/w:sectPr>/.exec(xml)?.[1] ?? '';
		expect(paragraphSection).toContain('<w:pgSz');
		expect(paragraphSection).not.toContain('pgBorders');
		expect(xml.match(/<w:pgBorders/g)).toHaveLength(1);
		const reopened = await loadDocx(bytes);
		expect(reopened.model.sections?.[0]?.pageBorders).toBeUndefined();
		expect(reopened.model.sections?.[1]?.pageBorders?.top?.style).toBe('single');
	});
});
