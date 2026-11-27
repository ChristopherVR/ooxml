import {
	assertMasterMarkup,
	masterAliases,
	masterCells,
	masterShapes,
	masterStyleResolver,
	overlay,
	type MasterCellSource,
} from './edit-master-index';

/** The effective cells a master template gives every instance on one style combination. */
export interface MasterTemplateLayer {
	readonly cells: ReadonlyMap<string, MasterCellSource>;
	readonly aliases: ReadonlyMap<Element, ReadonlySet<string>>;
}

/**
 * Builds the shared parts of master effective cells once per proof: a master's markup proof and
 * its template cells, and each template's style and template layer. Instances of one master then
 * add only their own cells, so the construction budget counts the cells built, not the same master
 * rebuilt for every instance on the page.
 */
export function masterLayers(document: Element, check: () => void, charge: () => void) {
	const styles = masterStyleResolver(document, check, charge);
	const templateCells = new Map<Element, ReadonlyMap<string, Element>>();
	const proven = new Map<Element, Element[]>();
	const layers = new Map<Element, Map<ReadonlyMap<string, Element>, MasterTemplateLayer>>();
	const cellsOf = (template: Element) => {
		let cells = templateCells.get(template);
		if (!cells) templateCells.set(template, (cells = masterCells(template, charge)));
		return cells;
	};
	return {
		styles,
		/** The master's sheets, after proving once that every formula in it is in a known cell. */
		templates(root: Element): readonly Element[] {
			let templates = proven.get(root);
			if (templates) return templates;
			templates = masterShapes(root);
			assertMasterMarkup(
				root,
				new Set(
					templates.flatMap((shape) => {
						charge();
						return [...cellsOf(shape).values()];
					}),
				),
				check,
			);
			proven.set(root, templates);
			return templates;
		},
		/** Style cells under the template's own cells, for an instance on `local`'s styles. */
		layer(template: Element, local: Element | undefined, moved: boolean): MasterTemplateLayer {
			const style = styles(template, local, moved);
			let byStyle = layers.get(template);
			if (!byStyle) layers.set(template, (byStyle = new Map()));
			let layer = byStyle.get(style);
			if (!layer) {
				const cells = new Map<string, MasterCellSource>();
				for (const [name, node] of style) cells.set(name, { node, kind: 'style' });
				const aliases = masterAliases(cells, charge);
				overlay(cells, cellsOf(template), 'template', charge, aliases);
				byStyle.set(style, (layer = { cells, aliases }));
			}
			return layer;
		},
	};
}
