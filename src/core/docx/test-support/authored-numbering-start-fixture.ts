import { numberingStartFixture, type NumberingStartCase } from './numbering-start-fixture';
import { loadDocx } from '../parse';
import { setListStartOverride } from '../numbering-start-edit';

/** Exercise the real core command/save path before an independent native Word reopen. */
export async function authoredNumberingStartFixture(item: NumberingStartCase): Promise<Uint8Array> {
	if (item.startOverride === null) throw new Error('An authored fixture requires a start value.');
	const loaded = await loadDocx(await numberingStartFixture({ ...item, startOverride: 7 }));
	return loaded.save({
		...loaded.model,
		numberingCatalog: setListStartOverride(
			loaded.model.numberingCatalog!,
			1,
			2,
			item.startOverride,
		),
	});
}
