/**
 * The theme part a page edit selects from. A built-in theme is written into the drawing the first
 * time one of its schemes (colours, fonts, effects or connectors) is picked; a part of the same
 * name written before font schemes and effect styles existed is replaced by the current one, since
 * without them the picked fonts and effects would not show.
 */
import { NS, parseXml } from '../xml/index';
import { RELATIONSHIP_TYPES } from '../opc/relationship-types';
import { relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import type { VisioPackage } from './package';
import { decodePath, fail } from './package-common';
import type { VisioTheme } from './theme';
import type { VisioBuiltInTheme } from './theme-builtins';
import { visioThemeXml } from './theme-write';

const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
const copy = (root: Element) => (root.ownerDocument!.cloneNode(true) as Document).documentElement;

/** Make sure the drawing holds the built-in theme `source`; new and replaced parts go to `created`. */
export async function ensureBuiltInThemePart(
	pkg: VisioPackage,
	parts: ReadonlyMap<string, Uint8Array>,
	documentPath: string,
	themes: readonly VisioTheme[],
	source: VisioBuiltInTheme,
	dirty: Map<string, Element>,
	created: Map<string, Uint8Array>,
	check: () => void,
): Promise<void> {
	const matches = themes.filter((theme) =>
		[theme.colorId, theme.effectId, theme.connectorId, theme.fontId].includes(source.schemeId),
	);
	if (matches.length > 1 || (matches[0] && matches[0].name !== source.name))
		fail(
			'EDIT_THEME_CONFLICT',
			'Another theme part already uses the scheme IDs of this built-in theme.',
		);
	if (matches[0]) {
		const current = matches[0];
		if (current.fontId === undefined && current.part && parts.has(current.part))
			created.set(current.part, new TextEncoder().encode(visioThemeXml(source)));
		return;
	}
	const relsPart = relationshipsPartFor(documentPath);
	const rels = parts.has(relsPart)
		? copy(await pkg.readXml(relsPart, 'Relationships'))
		: parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const types = copy(await pkg.readXml('[Content_Types].xml', 'Types'));
	const taken = new Set([...parts.keys()].map((name) => decodePath(name).toLowerCase()));
	const ids = new Set(
		Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship')).map(
			(node) => node.getAttribute('Id') ?? '',
		),
	);
	let index = 1;
	let target = '';
	let path = '';
	do {
		check();
		target = `theme/theme${index++}.xml`;
		path = resolvePartPath(documentPath, target);
	} while (taken.has(path.toLowerCase()));
	index = 1;
	while (ids.has(`rId${index}`)) index++;
	const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
	relationship.setAttribute('Id', `rId${index}`);
	relationship.setAttribute('Type', RELATIONSHIP_TYPES.theme);
	relationship.setAttribute('Target', target);
	rels.appendChild(relationship);
	const override = types.ownerDocument!.createElementNS(types.namespaceURI, 'Override');
	override.setAttribute('PartName', `/${path}`);
	override.setAttribute('ContentType', THEME_CONTENT_TYPE);
	types.appendChild(override);
	dirty.set(relsPart, rels);
	dirty.set('[Content_Types].xml', types);
	created.set(path, new TextEncoder().encode(visioThemeXml(source)));
}
