import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { loadBuiltinSmartArtStyleDefinitionXml } from '../../../diagram/builtin/style-definitions';
import { PptxHandler } from '../../core/PptxHandler';
import type { SmartArtPptxElement } from '../../core/types/elements';
import { readCorpusFixture } from './real-world-corpus-helpers';

const smartArtOf = (slides: { elements: unknown[] }[], id?: string) =>
	slides
		.flatMap((slide) => slide.elements as SmartArtPptxElement[])
		.find(
			(candidate) =>
				candidate.type === 'smartArt' &&
				(id ? candidate.id === id : Boolean(candidate.smartArtData?.quickStyle)),
		)!;

describe('built-in SmartArt colour and quick-style definitions on save', () => {
	it('writes the picked definitions as the parts and names them on the document point', async () => {
		const handler = new PptxHandler();
		const loaded = await handler.load(readCorpusFixture('smartart-chart-table-mix.pptx'));
		const element = smartArtOf(loaded.slides);
		const colors = loadBuiltinSmartArtStyleDefinitionXml('colors', 'accent4_3');
		const style = loadBuiltinSmartArtStyleDefinitionXml('style', '3d7');
		element.smartArtData!.builtinColorsXml = colors;
		element.smartArtData!.builtinQuickStyleXml = style;
		element.smartArtData!.style = 'moderate';
		element.smartArtData!.quickStyleDirty = true;
		const bytes = await handler.save(loaded.slides);
		const zip = await JSZip.loadAsync(bytes);
		const data = await Promise.all(
			Object.keys(zip.files)
				.filter((path) => /ppt\/diagrams\/data\d+\.xml$/u.test(path))
				.map((path) => zip.file(path)!.async('string')),
		);
		expect(
			data.some(
				(xml) =>
					xml.includes('csTypeId="urn:microsoft.com/office/officeart/2005/8/colors/accent4_3"') &&
					xml.includes('qsTypeId="urn:microsoft.com/office/officeart/2005/8/quickstyle/3d7"'),
			),
		).toBe(true);
		const reopened = await new PptxHandler().load(bytes);
		const saved = smartArtOf(reopened.slides, element.id).smartArtData!;
		expect(saved.colorTransform?.uniqueId).toBe(
			'urn:microsoft.com/office/officeart/2005/8/colors/accent4_3',
		);
		expect(saved.quickStyle?.uniqueId).toBe(
			'urn:microsoft.com/office/officeart/2005/8/quickstyle/3d7',
		);
		// Metallic Scene keeps its spatial camera and bevels in the saved quick style.
		expect(JSON.stringify(saved.quickStyle?.scene3d)).toContain('perspective');
	});
});
