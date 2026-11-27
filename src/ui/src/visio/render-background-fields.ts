import { visioBackgroundFieldShapes, type VisioDocument, type VisioPage } from 'ooxml-core/visio';

/**
 * The page's layers with each background page's page fields showing `page`, as Visio draws them:
 * PAGENAME() on a background reads the name of the foreground page that uses it.
 */
export function backgroundsUnder(
	model: VisioDocument,
	page: VisioPage,
	layers: VisioPage[],
): VisioPage[] {
	if (layers.length < 2) return layers;
	const index = model.pages.filter((item) => !item.isBackground).indexOf(page);
	const foreground = { name: page.name, ...(index >= 0 ? { number: index + 1 } : {}) };
	return layers.map((layer) => {
		if (layer === page) return layer;
		const shapes = visioBackgroundFieldShapes(layer, foreground);
		return shapes === layer.shapes ? layer : { ...layer, shapes };
	});
}
