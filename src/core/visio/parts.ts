import { relAttr } from '../xml/index.js';
import type { VisioConnection, VisioPage } from './model.js';
import { VisioPackage, VisioPackageError } from './package.js';
import { metadata } from './metadata.js';
import { attribute, child, children, VISIO_NS, VISIO_LEGACY_NS, type Report } from './sheet.js';
const REL = 'http://schemas.microsoft.com/visio/2010/relationships/';
export async function visioXml(
	pkg: VisioPackage,
	path: string,
	rootName: string,
): Promise<Element> {
	const root = await pkg.readXml(path, rootName);
	if (root.namespaceURI !== VISIO_NS && root.namespaceURI !== VISIO_LEGACY_NS)
		throw new VisioPackageError('INVALID_VISIO_XML', `Unexpected namespace in ${path}.`);
	return root;
}
export async function related(
	pkg: VisioPackage,
	source: string,
	type: string,
	required = true,
): Promise<string | undefined> {
	const matches = [...(await pkg.relationships(source)).values()].filter(
		(rel) => rel.type === REL + type,
	);
	if (
		matches.length > 1 ||
		(required && matches.length !== 1) ||
		matches.some((r) => r.mode !== 'Internal')
	)
		throw new VisioPackageError(
			'INVALID_RELATIONSHIP',
			`Expected one internal ${type} relationship from ${source || 'package root'}.`,
		);
	return matches[0]?.target;
}
export async function indexedPart(
	pkg: VisioPackage,
	source: string,
	node: Element,
	type: string,
): Promise<string> {
	const id = relAttr(child(node, 'Rel'), 'id');
	if (id !== undefined) metadata(id, 256, 'Relationship ID');
	const rel = id ? (await pkg.relationships(source)).get(id) : undefined;
	if (!rel || rel.mode !== 'Internal' || rel.type !== REL + type)
		throw new VisioPackageError(
			'INVALID_RELATIONSHIP',
			`Invalid ${type} relationship ${id ?? '(missing)'} in ${source}.`,
		);
	return rel.target;
}
export function connections(root: Element, report: Report): VisioConnection[] {
	return children(child(root, 'Connects'), 'Connect').flatMap((node) => {
		const fromShapeId = attribute(node, 'FromSheet'),
			toShapeId = attribute(node, 'ToSheet');
		if (!fromShapeId || !toShapeId) {
			report('invalid-connection', 'Connection without both shape IDs was omitted.');
			return [];
		}
		const from = Number(attribute(node, 'FromPart')),
			to = Number(attribute(node, 'ToPart'));
		return [
			{
				fromShapeId: metadata(fromShapeId, 256, 'Connection shape ID'),
				toShapeId: metadata(toShapeId, 256, 'Connection shape ID'),
				fromCell: metadata(attribute(node, 'FromCell') ?? '', 256, 'Connection cell'),
				toCell: metadata(attribute(node, 'ToCell') ?? '', 256, 'Connection cell'),
				...(Number.isSafeInteger(from) ? { fromPart: from } : {}),
				...(Number.isSafeInteger(to) ? { toPart: to } : {}),
			},
		];
	});
}
export function validateBackgrounds(pages: VisioPage[], report: Report): void {
	const byId = new Map(pages.map((page) => [page.id, page]));
	for (const page of pages) {
		const seen = new Set<string>([page.id]);
		let background = page.backgroundPageId;
		while (background !== undefined) {
			if (seen.has(background)) {
				report('background-cycle', 'Cyclic background page reference was omitted.', {
					pageId: page.id,
				});
				delete page.backgroundPageId;
				break;
			}
			seen.add(background);
			const next = byId.get(background);
			if (!next) {
				report('missing-background-page', `Background page ${background} could not be resolved.`, {
					pageId: page.id,
				});
				delete page.backgroundPageId;
				break;
			}
			background = next.backgroundPageId;
		}
	}
}
