import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { afterEach, expect, it, vi } from 'vitest';

import { PptxHandlerRuntime } from './PptxHandlerRuntimeState';

const partPath = 'ppt/slides/slide1.xml';
const relsPath = 'ppt/slides/_rels/slide1.xml.rels';

class RelationshipRuntime extends PptxHandlerRuntime {
	constructor(zip: JSZip) {
		super();
		this.zip = zip;
		this.parser = new XMLParser({ ignoreAttributes: false });
	}

	async relationships() {
		await this.loadPartRelationships(partPath);
		return {
			targets: this.slideRelsMap.get(partPath),
			externalIds: this.externalRelsMap.get(partPath),
		};
	}
}

afterEach(() => {
	vi.restoreAllMocks();
});

it('reads relationships once without a load cache and preserves internal and external targets', async () => {
	const bytes = await new JSZip()
		.file(
			relsPath,
			'<Relationships><Relationship Id="rId1" Target="../slideLayouts/layout1.xml"/>' +
				'<Relationship Id="rId2" Target="https://example.com" TargetMode="External"/></Relationships>',
		)
		.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
	const zip = await JSZip.loadAsync(bytes);
	const reads = vi.spyOn(zip.file(relsPath)!, 'async');
	const result = await new RelationshipRuntime(zip).relationships();
	expect(reads).toHaveBeenCalledTimes(1);
	expect(result.targets).toEqual(
		new Map([
			['rId1', '../slideLayouts/layout1.xml'],
			['rId2', 'https://example.com'],
		]),
	);
	expect(result.externalIds).toEqual(new Set(['rId2']));
});

it.each(['missing', 'empty'])('ignores a %s relationship part', async (kind) => {
	const zip = new JSZip();
	if (kind === 'empty') zip.file(relsPath, '');
	expect(await new RelationshipRuntime(zip).relationships()).toEqual({
		targets: undefined,
		externalIds: undefined,
	});
});
