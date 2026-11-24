// The stories Find searches beside the body (`findStories` gives Word's order), each turned into a
// ProseMirror document the same way its in-place editor builds one, so a match's positions are
// the positions in that editor. Text boxes are image nodes of the body; each is a unit of the Text
// Boxes story, its lines a paragraph each.
import { EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	createDocument,
	type Block,
	type DocumentModel,
	type ReplaceFormatPatch,
} from 'ooxml-core/docx';
import {
	findStories,
	findStoryGroup,
	findStoryKey,
	type FindOptions,
	type FindStory,
} from 'ooxml-core/docx/ui';
import { docToModel, modelToDoc } from './model-adapter';
import { withMark, withoutMark } from './note-editor';
import { findDocumentMatches, replaceTargets, type DocumentMatch } from './search-document';
import type { ReplacementTarget } from './search-document';
import { schema } from './schema';

/** One searchable unit: a story, or one text box of the Text Boxes story. */
export interface SearchUnit {
	readonly story: FindStory;
	/** The unit's identity: the story key, or `textBox:<pos>` for a text box. */
	readonly key: string;
	/** Replace All applies one undo step per group (see `findStoryGroup`). */
	readonly group: string;
	readonly doc: ProseMirrorNode;
	/** The text box's position in the body (text boxes only). */
	readonly textBoxPos?: number;
	/** A text box Word wrote in a form this editor can rewrite; other units are always true. */
	readonly replaceable: boolean;
}

export interface StoryMatch extends DocumentMatch {
	readonly unit: SearchUnit;
	/** The unit's index in search order. */
	readonly order: number;
}

/** An open in-place story editor, whose live document stands in for the model's copy. */
export interface OpenStory {
	readonly key: string;
	readonly doc: ProseMirrorNode;
}

const docs = new WeakMap<readonly Block[], ProseMirrorNode>();

/** The editable blocks of a story: a note loses its automatic number mark, as its editor does. */
function editableBlocks(story: FindStory): Block[] {
	if (story.kind === 'footnote' || story.kind === 'endnote')
		return withoutMark([...story.blocks]).editable;
	return story.kind === 'header' || story.kind === 'footer' ? [...story.blocks] : [];
}

function storyDoc(story: FindStory & { blocks: readonly Block[] }): ProseMirrorNode {
	const cached = docs.get(story.blocks);
	if (cached) return cached;
	const doc = modelToDoc({ ...createDocument(), blocks: editableBlocks(story) });
	docs.set(story.blocks, doc);
	return doc;
}

function textBoxLines(value: unknown): string[] {
	if (typeof value !== 'string') return [];
	try {
		const parsed: unknown = JSON.parse(value);
		return Array.isArray(parsed) ? parsed.map(String) : [];
	} catch {
		return [];
	}
}

function linesDoc(lines: readonly string[]): ProseMirrorNode {
	const paragraph = schema.nodes.paragraph;
	return schema.node(
		'doc',
		null,
		(lines.length ? lines : ['']).map((line) =>
			paragraph.create(null, line ? schema.text(line) : null),
		),
	);
}

function textBoxUnits(body: ProseMirrorNode): SearchUnit[] {
	const story: FindStory = { kind: 'textBoxes' };
	const units: SearchUnit[] = [];
	body.descendants((node, pos) => {
		if (node.type.name !== 'image' || typeof node.attrs.textBoxText !== 'string') return true;
		units.push({
			story,
			key: `textBox:${pos}`,
			group: findStoryGroup(story),
			doc: linesDoc(textBoxLines(node.attrs.textBoxText)),
			textBoxPos: pos,
			replaceable: Boolean(node.attrs.textBoxEditable),
		});
		return false;
	});
	return units;
}

/** The body only: Find without a story host (the surface mode's host dialog). */
export function mainUnits(body: ProseMirrorNode): SearchUnit[] {
	const story: FindStory = { kind: 'main' };
	return [{ story, key: 'main', group: 'main', doc: body, replaceable: true }];
}

/** Every unit of the document in Word's order. `open` replaces its story's model copy. */
export function searchUnits(
	body: ProseMirrorNode,
	model: DocumentModel,
	open?: OpenStory,
): SearchUnit[] {
	return findStories(model).flatMap((story): SearchUnit[] => {
		if (story.kind === 'main') return mainUnits(body);
		if (story.kind === 'textBoxes') return textBoxUnits(body);
		const key = findStoryKey(story);
		const doc = open?.key === key ? open.doc : storyDoc(story);
		return [{ story, key, group: findStoryGroup(story), doc, replaceable: true }];
	});
}

/** Every match in every unit, in search order. */
export function findUnitMatches(
	units: readonly SearchUnit[],
	query: string,
	options: FindOptions,
): StoryMatch[] {
	return units.flatMap((unit, order) =>
		findDocumentMatches(unit.doc, query, options).map((match) => ({ ...match, unit, order })),
	);
}

/** A header, footer or note's blocks after the replacements (in its editor's positions). */
export function replacedStoryBlocks(
	unit: SearchUnit,
	targets: readonly ReplacementTarget[],
	replaceFormat?: ReplaceFormatPatch,
): Block[] {
	const { story } = unit;
	const original = editableBlocks(story);
	const doc = replaceTargets(EditorState.create({ doc: unit.doc }), targets, replaceFormat).doc;
	const blocks = docToModel(doc, { ...createDocument(), blocks: original }).blocks;
	if (story.kind !== 'footnote' && story.kind !== 'endnote') return blocks;
	return withMark(blocks, withoutMark([...story.blocks]).mark);
}

/** A text box's lines after the replacements; a paragraph mark in a replacement starts a line. */
export function replacedTextBoxLines(
	unit: SearchUnit,
	targets: readonly ReplacementTarget[],
): string[] {
	const doc = replaceTargets(EditorState.create({ doc: unit.doc }), targets).doc;
	const lines: string[] = [];
	doc.forEach((paragraph) => lines.push(paragraph.textContent));
	return lines;
}
