import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../index';
import { mapCustomShowActions } from '../../core/utils/custom-show-actions';
import { pptxTools } from '../catalogue/index';
import { NS } from '../../../xml/namespaces';
import { elements, parseXml } from '../../../xml/xml';

// One real deck with custom show 7 and shape, text-run and table-cell links into it.
const fixture = new Uint8Array(
	readFileSync(new URL('../../__tests__/fixtures/custom-show-hyperlinks.pptx', import.meta.url)),
);

async function actions(bytes: Uint8Array): Promise<string[]> {
	const zip = await JSZip.loadAsync(bytes);
	const result: string[] = [];
	for (const path of Object.keys(zip.files).filter((name) =>
		/^ppt\/slides\/slide\d+\.xml$/.test(name),
	)) {
		const doc = parseXml(await zip.file(path)!.async('string'));
		for (const name of ['hlinkClick', 'hlinkHover', 'hlinkMouseOver']) {
			for (const node of Array.from(doc.getElementsByTagNameNS(NS.a, name))) {
				const action = node.getAttribute('action');
				if (action) result.push(action);
			}
		}
	}
	return result;
}

async function tableContent(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const doc = parseXml(await zip.file('ppt/slides/slide1.xml')!.async('string'));
	const table = doc.getElementsByTagNameNS(NS.a, 'tbl')[0]!;
	const paragraph = table.getElementsByTagNameNS(NS.a, 'p')[0]!;
	return elements(paragraph)
		.filter((node) => ['r', 'fld', 'br'].includes(node.localName))
		.map((node) => {
			const props = elements(node).find((child) => child.localName === 'rPr');
			return {
				kind: node.localName,
				text: node.textContent,
				id: node.getAttribute('id'),
				type: node.getAttribute('type'),
				bold: props?.getAttribute('b'),
				italic: props?.getAttribute('i'),
			};
		});
}

describe('custom show hyperlink round trips', () => {
	it('deletes all three jumps through the registered tool and saves the cleared selection', async () => {
		expect(await actions(fixture)).toHaveLength(3);
		const tool = pptxTools.find((candidate) => candidate.name === 'manage_custom_shows');
		if (tool?.kind !== 'edit') throw new Error('Missing custom show edit tool');
		const result = await tool.run(fixture, { action: 'delete', showId: '7' });
		if (!result.bytes) throw new Error('The edit did not save');
		expect(result.clearedActions).toBe(3);
		expect(await actions(result.bytes)).toEqual([]);
		expect(await tableContent(result.bytes)).toEqual(await tableContent(fixture));
		const handler = new PptxHandler();
		try {
			const loaded = await handler.load(new Uint8Array(result.bytes).buffer);
			expect(loaded.customShows ?? []).toEqual([]);
			expect(loaded.presentationProperties?.showSlidesMode).toBe('all');
		} finally {
			handler.dispose();
		}
	});

	it('keeps all three jumps in step when host ids are normalized on save', async () => {
		const handler = new PptxHandler();
		try {
			const loaded = await handler.load(new Uint8Array(fixture).buffer);
			const show = loaded.customShows![0]!;
			show.id = 'host-show';
			loaded.presentationProperties!.showSlidesCustomShowId = 'host-show';
			const original = structuredClone(loaded.slides);
			const mapped = mapCustomShowActions(loaded.slides, (id) => (id === '7' ? 'host-show' : id));
			expect(loaded.slides).toStrictEqual(original);
			expect(mapped.changed).toBe(3);
			const bytes = await handler.save(mapped.slides, {
				customShows: loaded.customShows,
				presentationProperties: loaded.presentationProperties,
			});
			expect(await actions(bytes)).toEqual(Array(3).fill('ppaction://customshow?id=0&return=true'));
			expect(await tableContent(bytes)).toEqual(await tableContent(fixture));
			const reader = new PptxHandler();
			try {
				const saved = await reader.load(new Uint8Array(bytes).buffer);
				expect(saved.customShows?.[0]?.id).toBe('0');
				expect(saved.presentationProperties?.showSlidesCustomShowId).toBe('0');
			} finally {
				reader.dispose();
			}
		} finally {
			handler.dispose();
		}
	});
});
