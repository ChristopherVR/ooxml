import { describe, expect, it } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import type { XmlObject } from '../types';
import { customGeometryPathsToXml } from './custom-geometry';
import {
	orderedPathCommandEntries,
	readPathListCommandOrder,
	restorePathListCommandOrder,
} from './custom-geometry-command-order';
import { parseStructuredCustomGeometry } from './custom-geometry-parser';

const ensureArray = (value: unknown): unknown[] =>
	value === undefined ? [] : Array.isArray(value) ? value : [value];

function customGeometryFrom(parsed: XmlObject): XmlObject {
	return ((parsed['p:sp'] as XmlObject)['p:spPr'] as XmlObject)['a:custGeom'] as XmlObject;
}

describe('custom geometry command order', () => {
	it('preserves interleaved commands through parse, model, serialize, and reload', () => {
		const xml =
			'<p:sp xmlns:p="urn:p" xmlns:a="urn:a"><p:spPr><a:custGeom>' +
			'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
			'<a:pathLst><a:path w="100" h="100">' +
			'<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
			'<a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
			'<a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>' +
			'<a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/>' +
			'</a:path></a:pathLst></a:custGeom></p:spPr></p:sp>';
		const factory = new PptxRuntimeDependencyFactory();
		const parser = factory.createParser();
		const parsed = parser.parse(xml) as XmlObject;
		const paths = parseStructuredCustomGeometry(customGeometryFrom(parsed), 100, 100, ensureArray);

		expect(paths[0].segments.map((segment) => segment.type)).toStrictEqual([
			'moveTo',
			'lineTo',
			'arcTo',
			'lineTo',
			'close',
		]);

		const serialized = factory.createBuilder().build({
			'p:sp': {
				'@_xmlns:p': 'urn:p',
				'@_xmlns:a': 'urn:a',
				'p:spPr': { 'a:custGeom': customGeometryPathsToXml(paths) },
			},
		});
		const serializedPath = /<a:path\b[^>]*>([\s\S]*?)<\/a:path>/u.exec(serialized)?.[1] ?? '';
		expect(
			[...serializedPath.matchAll(/<a:(moveTo|lnTo|arcTo|close)\b/gu)].map((match) => match[1]),
		).toStrictEqual(['moveTo', 'lnTo', 'arcTo', 'lnTo', 'close']);

		const reparsed = parser.parse(serialized) as XmlObject;
		const reparsedPaths = parseStructuredCustomGeometry(
			customGeometryFrom(reparsed),
			100,
			100,
			ensureArray,
		);
		expect(reparsedPaths[0].segments.map((segment) => segment.type)).toStrictEqual([
			'moveTo',
			'lineTo',
			'arcTo',
			'lineTo',
			'close',
		]);
	});

	describe('order as plain data (survives a copy of the tree)', () => {
		const xml =
			'<p:sp xmlns:p="urn:p" xmlns:a="urn:a"><p:spPr><a:custGeom><a:pathLst>' +
			'<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
			'<a:cubicBezTo><a:pt x="1" y="1"/><a:pt x="2" y="2"/><a:pt x="3" y="3"/></a:cubicBezTo>' +
			'<a:lnTo><a:pt x="4" y="4"/></a:lnTo>' +
			'<a:cubicBezTo><a:pt x="5" y="5"/><a:pt x="6" y="6"/><a:pt x="7" y="7"/></a:cubicBezTo>' +
			'<a:close/></a:path>' +
			'<a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
			'<a:lnTo><a:pt x="10" y="0"/></a:lnTo></a:path>' +
			'</a:pathLst></a:custGeom></p:spPr></p:sp>';
		const order = ['moveTo', 'cubicBezTo', 'lnTo', 'cubicBezTo', 'close'];
		const names = (path: XmlObject): string[] =>
			orderedPathCommandEntries(path, ensureArray).map(([key]) => key.replace('a:', ''));
		const pathLstOf = (parsed: XmlObject): XmlObject =>
			customGeometryFrom(parsed)['a:pathLst'] as XmlObject;
		const firstPath = (pathLst: XmlObject): XmlObject =>
			ensureArray(pathLst['a:path'])[0] as XmlObject;

		it('loses the order on a copy, and restores it from the recorded data', () => {
			const parsed = new PptxRuntimeDependencyFactory().createParser().parse(xml) as XmlObject;
			const pathLst = pathLstOf(parsed);
			expect(names(firstPath(pathLst))).toStrictEqual(order);

			const recorded = readPathListCommandOrder(pathLst);
			expect(recorded).toStrictEqual([order, ['moveTo', 'lnTo']]);

			const copy = structuredClone(pathLst);
			// The copy's objects were never annotated: grouped by type.
			expect(names(firstPath(copy))).toStrictEqual([
				'moveTo',
				'cubicBezTo',
				'cubicBezTo',
				'lnTo',
				'close',
			]);
			restorePathListCommandOrder(copy, structuredClone(recorded));
			expect(names(firstPath(copy))).toStrictEqual(order);
		});

		it('ignores an order whose commands do not match the path', () => {
			const parsed = new PptxRuntimeDependencyFactory().createParser().parse(xml) as XmlObject;
			const copy = structuredClone(pathLstOf(parsed));
			restorePathListCommandOrder(copy, [['moveTo', 'lnTo', 'close']]);
			expect(names(firstPath(copy))).toStrictEqual([
				'moveTo',
				'cubicBezTo',
				'cubicBezTo',
				'lnTo',
				'close',
			]);
		});

		it('returns undefined when no path carries an order', () => {
			expect(readPathListCommandOrder(undefined)).toBeUndefined();
			expect(readPathListCommandOrder({ 'a:path': { 'a:moveTo': {} } })).toBeUndefined();
		});
	});
});
