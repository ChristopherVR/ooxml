import { describe, expect, it } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { projectChartCompatibilityTree } from '../../core/utils/chart-compat-tree';
import { readOriginalChartPart } from '../../core/utils/chart-part-repair';
import { classicChartDecks } from './chart-save-parity';

/**
 * The compatibility tree projected from the shared chart DOM stands in for the object parser on
 * load. Over every classic chart part of every committed deck, a projection that is produced must
 * be exactly the object parser's tree; parts it defers to the parser are only counted.
 */
describe('chart compatibility tree over the committed decks', () => {
	const parser = new PptxRuntimeDependencyFactory().createParser();

	it('matches the object parser wherever it projects a part', async () => {
		let projectedParts = 0;
		let parts = 0;
		for (const deck of await classicChartDecks())
			for (const { part } of deck.parts) {
				const xml = await deck.archive.file(part)!.async('string');
				parts++;
				const projected = projectChartCompatibilityTree(parser, xml, readOriginalChartPart(xml));
				if (projected === undefined) continue;
				projectedParts++;
				// The readers look the root up by name; only the declaration entry may be absent.
				const { '?xml': _declaration, ...parsed } = parser.parse(xml) as Record<string, unknown>;
				expect(JSON.stringify(projected), `${deck.name} ${part}`).toBe(JSON.stringify(parsed));
			}
		expect(parts).toBeGreaterThan(0);
		// The projection must cover a real share of the corpus, or it is not doing its job.
		expect(projectedParts).toBeGreaterThan(0);
	});
});
