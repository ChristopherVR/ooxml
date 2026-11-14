import { describe, expect, it } from 'vitest';
import { createDocument, type Block, type DocumentModel } from './model';
import type { HeaderFooterContent, SectionProperties } from './section-model';
import { findStories, findStoryGroup, findStoryKey } from './find-stories';

const blocks = (text: string): Block[] => [{ type: 'paragraph', id: text, runs: [{ text }] }];
const part = (text: string, partName?: string): HeaderFooterContent => ({
	blocks: blocks(text),
	...(partName ? { partName } : {}),
});
const section = (extra: Partial<SectionProperties>): SectionProperties =>
	({ ...createDocument().sections?.[0], ...extra }) as SectionProperties;

describe('findStories', () => {
	it('lists the stories in the order Word Find runs through them', () => {
		const shared = part('shared header', 'word/header1.xml');
		const model: DocumentModel = {
			...createDocument(),
			footnotes: [{ id: '1', blocks: blocks('note one') }],
			endnotes: [{ id: '2', blocks: blocks('end one') }],
			sections: [
				section({
					headers: { default: shared, first: part('first header'), even: part('even header') },
					footers: { default: part('footer') },
				}),
				section({
					headers: { default: { ...shared } },
					footers: { first: part('second first footer') },
				}),
			],
		};
		expect(findStories(model).map(findStoryKey)).toEqual([
			'main',
			'header:0:even',
			'header:0:default',
			'footer:0:default',
			'header:0:first',
			'footer:1:first',
			'footnote:1',
			'endnote:2',
			'textBoxes',
		]);
	});

	it('groups every note of a kind into one undo story', () => {
		const stories = findStories({
			...createDocument(),
			footnotes: [
				{ id: '1', blocks: blocks('a') },
				{ id: '2', blocks: blocks('b') },
			],
			sections: [section({ headers: { default: part('h') } })],
		});
		expect(stories.map(findStoryGroup)).toEqual([
			'main',
			'header:0:default',
			'footnote',
			'footnote',
			'textBoxes',
		]);
	});
});
