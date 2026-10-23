import { beforeAll, describe, expect, it } from 'vitest';

import {
	diagramParagraphsText,
	projectDiagramNodeTextNeutral,
	readDiagramNodeParagraphs,
	type DiagramNeutralTextSegment,
} from '../../../diagram/index';
import {
	fixtureTextBodies,
	type FixtureTextBody,
} from '../../../drawingml/test-support/text-fixtures';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { parseSmartArtTextParagraphs } from '../../core/core/runtime/smartart-text-paragraphs';
import type { TextSegment, XmlObject } from '../../core/types';
import { projectSmartArtNodeText } from '../../core/utils/smartart-node-text-projection';

/**
 * The shared node-text projection (`diagram/text/node-text-projection.ts`) driven by the pptx
 * styler and by the neutral styler must build the same segment structure (texts, breaks, levels,
 * bullets) for every SmartArt data point of the committed fixtures, so Word and Excel get the
 * node text PowerPoint shows.
 */

const parser = new PptxRuntimeDependencyFactory().createParser();

function points(tree: unknown, into = new Map<string, XmlObject>()): Map<string, XmlObject> {
	if (!tree || typeof tree !== 'object') return into;
	for (const node of Array.isArray(tree) ? tree : [tree]) {
		if (!node || typeof node !== 'object') continue;
		const object = node as XmlObject;
		if (object['@_modelId'] !== undefined) into.set(String(object['@_modelId']), object);
		for (const [key, value] of Object.entries(object))
			if (!key.startsWith('@_')) points(value, into);
	}
	return into;
}

function pptxBullet(segment: TextSegment): string | undefined {
	const bullet = segment.bulletInfo;
	if (!bullet) return undefined;
	if (bullet.none) return 'none';
	return bullet.char !== undefined ? `char:${bullet.char}` : `auto:${bullet.autoNumType}`;
}

function neutralBullet(segment: DiagramNeutralTextSegment): string | undefined {
	const bullet = segment.bulletInfo;
	if (!bullet) return undefined;
	if (bullet.kind === 'none') return 'none';
	return bullet.kind === 'char'
		? `char:${bullet.char}`
		: bullet.kind === 'autoNumber'
			? `auto:${bullet.scheme}`
			: 'picture';
}

function structure(
	segments: (TextSegment | DiagramNeutralTextSegment)[],
	bullet: (segment: never) => string | undefined,
): unknown[] {
	return segments.map((segment) => [
		segment.text,
		segment.isParagraphBreak ?? false,
		segment.isLineBreak ?? false,
		segment.paragraphLevel,
		segment.fieldType,
		bullet(segment as never),
	]);
}

describe('SmartArt node text: the pptx and neutral stylers project the same structure', () => {
	let bodies: FixtureTextBody[] = [];
	beforeAll(async () => {
		bodies = (await fixtureTextBodies()).filter((body) => body.element.localName === 't');
	}, 120_000);

	it('agrees for every data point text of the committed SmartArt parts', () => {
		const trees = new Map<string, Map<string, XmlObject>>();
		const differences: string[] = [];
		let compared = 0;
		for (const body of bodies) {
			const key = `${body.file}!${body.part}`;
			if (!trees.has(key)) trees.set(key, points(parser.parse(body.partXml)));
			const modelId = (body.element.parentNode as Element).getAttribute('modelId') ?? '';
			const point = trees.get(key)!.get(modelId);
			const pptxParagraphs = point ? parseSmartArtTextParagraphs(point) : undefined;
			const neutralParagraphs = readDiagramNodeParagraphs(body.element);
			if (!pptxParagraphs || !neutralParagraphs) {
				if (pptxParagraphs || neutralParagraphs) differences.push(`${body.label}: one side empty`);
				continue;
			}
			const text = diagramParagraphsText(neutralParagraphs);
			expect(diagramParagraphsText(pptxParagraphs), body.label).toBe(text);
			for (const bulletEnabled of [false, true]) {
				const pptx = projectSmartArtNodeText(
					{ id: modelId, text, paragraphs: pptxParagraphs },
					{},
					{ bulletEnabled },
				);
				const neutral = projectDiagramNodeTextNeutral(
					{ text, paragraphs: neutralParagraphs },
					undefined,
					{ bulletEnabled },
				);
				compared += pptx.length;
				if (
					JSON.stringify(structure(pptx, pptxBullet)) !==
					JSON.stringify(structure(neutral, neutralBullet))
				)
					differences.push(body.label);
			}
		}
		expect(differences).toEqual([]);
		expect(compared).toBeGreaterThan(1000);
	});

	it('reconciles an edited flat text the same way on both models', () => {
		const body = bodies.find((entry) => (entry.element.textContent ?? '').length > 3)!;
		const modelId = (body.element.parentNode as Element).getAttribute('modelId') ?? '';
		const point = points(parser.parse(body.partXml)).get(modelId)!;
		const edited = `Edited ${body.element.textContent}\nSecond line`;
		const pptx = projectSmartArtNodeText({
			id: modelId,
			text: edited,
			paragraphs: parseSmartArtTextParagraphs(point)!,
		});
		const neutral = projectDiagramNodeTextNeutral({
			text: edited,
			paragraphs: readDiagramNodeParagraphs(body.element)!,
		});
		expect(structure(neutral, neutralBullet)).toEqual(structure(pptx, pptxBullet));
		expect(neutral.map((segment) => segment.text).join('')).toContain('Second line');
	});
});
