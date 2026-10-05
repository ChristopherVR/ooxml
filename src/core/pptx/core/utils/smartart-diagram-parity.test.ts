// The pptx object-tree parsers and the neutral DOM parsers of the `diagram` area must agree on real
// parts: pptx re-imports the attribute parsers, the data-model parser is a separate implementation.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseDiagramDataModel } from '../../../diagram/index.js';
import type { XmlObject } from '../types';
import { collectSmartArtTransitionText } from './smartart-connector-labels';
import {
	parseSmartArtConnection,
	parseSmartArtPointCustomLayout,
} from './smartart-data-model-attributes';

const fixtures = path.join(import.meta.dirname, '../../__tests__/fixtures');
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
const list = (value: unknown): XmlObject[] =>
	value === undefined ? [] : ((Array.isArray(value) ? value : [value]) as XmlObject[]);

const DECKS = [
	['corpus', 'smartart-orgchart-assistants.pptx', 'ppt/diagrams/data2.xml'],
	['corpus', 'smartart-orgchart-many.pptx', 'ppt/diagrams/data1.xml'],
	['e2e', 'smartart-build-reveal.pptx', 'ppt/diagrams/data1.xml'],
] as const;

describe.each(DECKS)('diagram parity on %s/%s', (folder, deck, part) => {
	it('parses the same connections, transition labels and custom layouts', async () => {
		const zip = await JSZip.loadAsync(readFileSync(path.join(fixtures, folder, deck)));
		const xml = await zip.file(part)!.async('string');
		const tree = parser.parse(xml) as XmlObject;
		const dataModel = tree['dgm:dataModel'] as XmlObject;
		const points = list((dataModel['dgm:ptLst'] as XmlObject)['dgm:pt']);
		const cxns = list((dataModel['dgm:cxnLst'] as XmlObject)['dgm:cxn']);
		const neutral = parseDiagramDataModel(xml);

		const legacy = cxns.flatMap((cxn) => parseSmartArtConnection(cxn) ?? []);
		const labels = collectSmartArtTransitionText(points, (point) =>
			list((point['dgm:t'] as XmlObject | undefined)?.['a:p'])
				.flatMap((paragraph) => list(paragraph['a:r']))
				.map((run) => String(run['a:t'] ?? ''))
				.join(''),
		);
		const withLabels = legacy.map((connection) => {
			const label =
				(connection.parentTransitionId && labels.get(connection.parentTransitionId)) ||
				(connection.siblingTransitionId && labels.get(connection.siblingTransitionId)) ||
				undefined;
			return label ? { ...connection, label } : connection;
		});
		expect(neutral.connections).toEqual(withLabels);
		expect(neutral.connections.length).toBeGreaterThan(0);

		const legacyCustom = points.flatMap((point) => {
			const custom = parseSmartArtPointCustomLayout(point['dgm:prSet'] as XmlObject | undefined);
			return custom ? [[String(point['@_modelId']), custom] as const] : [];
		});
		const neutralCustom = neutral.points.flatMap((point) =>
			point.customLayout ? [[point.modelId, point.customLayout] as const] : [],
		);
		expect(neutralCustom).toEqual(legacyCustom);
	});
});
