import { NS } from '../xml/index';
import { VisioPackage } from './package';
import { fail, type VisioPackageLimits } from './package-common';
import { indexedPart, related, visioXml } from './parts';
import { attribute, child, children } from './sheet';
import { addMasterToDrawing, copy, type EnsuredStencilMaster } from './edit-master-add';

const STYLE_ATTRIBUTES = ['LineStyle', 'FillStyle', 'TextStyle'] as const;
const MAX_STYLE_DEPTH = 8;
const STENCIL_TYPES = [
	'application/vnd.ms-visio.stencil.main+xml',
	'application/vnd.ms-visio.template.main+xml',
	'application/vnd.ms-visio.drawing.main+xml',
];

/**
 * Brings the style sheets a stencil master uses into the drawing. A style the drawing has by name
 * is reused; any other is copied with a new ID, with the styles it is based on. Style 0 stays 0.
 */
function styleMapper(stencilDocument: Element, document: Element, touch: () => void) {
	const sheets = (root: Element) => children(child(root, 'StyleSheets'), 'StyleSheet');
	const source = new Map(sheets(stencilDocument).map((sheet) => [attribute(sheet, 'ID'), sheet]));
	const nameOf = (sheet: Element) => attribute(sheet, 'NameU') ?? attribute(sheet, 'Name');
	const done = new Map<string, string>();
	const map = (id: string | undefined, depth = 0): string | undefined => {
		if (id === undefined) return undefined;
		const known = done.get(id);
		if (known !== undefined) return known;
		const sheet = source.get(id);
		if (id === '0' || !sheet || depth > MAX_STYLE_DEPTH) return '0';
		const name = nameOf(sheet);
		const existing = sheets(document).find((candidate) => nameOf(candidate) === name);
		if (existing) {
			done.set(id, attribute(existing, 'ID') ?? '0');
			return done.get(id)!;
		}
		const ids = sheets(document).map((candidate) => Number(attribute(candidate, 'ID')));
		const created = String(Math.max(0, ...ids.filter(Number.isFinite)) + 1);
		done.set(id, created);
		const clone = document.ownerDocument!.importNode(sheet, true) as Element;
		clone.setAttribute('ID', created);
		for (const based of STYLE_ATTRIBUTES) {
			const mapped = map(attribute(sheet, based), depth + 1);
			if (mapped !== undefined) clone.setAttribute(based, mapped);
		}
		let container = child(document, 'StyleSheets');
		if (!container) {
			container = document.ownerDocument!.createElementNS(document.namespaceURI, 'StyleSheets');
			document.appendChild(container);
		}
		container.appendChild(clone);
		touch();
		return created;
	};
	return (node: Element) => {
		for (const name of STYLE_ATTRIBUTES) {
			const mapped = map(attribute(node, name));
			if (mapped !== undefined) node.setAttribute(name, mapped);
		}
	};
}

const IMAGE_RELATIONSHIP =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
/** The pictures a master part owns; any other related part is refused rather than dropped. */
async function masterPictures(pkg: VisioPackage, part: string) {
	const media: { relId: string; extension: string; bytes: Uint8Array }[] = [];
	for (const rel of (await pkg.relationships(part)).values()) {
		const extension = /\.([A-Za-z0-9]{1,5})$/.exec(rel.target)?.[1];
		if (rel.type !== IMAGE_RELATIONSHIP || rel.mode !== 'Internal' || !extension)
			fail(
				'UNSUPPORTED_EDIT_PACKAGE',
				'A stencil master that owns parts other than pictures cannot be copied into a drawing yet.',
			);
		media.push({ relId: rel.id, extension, bytes: await pkg.readBytes(rel.target) });
	}
	return media;
}

/**
 * Copy one master of a stencil package (.vssx, .vstx or a drawing) into the drawing's masters
 * part, with its styles, as Visio does when a shape of an opened stencil is first dropped. A copy
 * made earlier is found by its UniqueID and reused. A master that owns other parts (pictures)
 * other than pictures is refused rather than copied without them.
 */
export async function importStencilMaster(
	source: Uint8Array,
	stencil: Uint8Array,
	stencilMasterId: string,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EnsuredStencilMaster> {
	const pkg = await VisioPackage.open(stencil, limits);
	const types = await pkg.readXml('[Content_Types].xml', 'Types');
	const documentPart = await related(pkg, '', 'document');
	const override = documentPart
		? Array.from(types.getElementsByTagNameNS(types.namespaceURI, 'Override')).find(
				(node) => node.getAttribute('PartName') === `/${documentPart}`,
			)
		: undefined;
	if (!documentPart || !STENCIL_TYPES.includes(override?.getAttribute('ContentType') ?? ''))
		fail('UNSUPPORTED_FORMAT', 'The stencil is not a Visio stencil package.');
	const stencilDocument = await visioXml(pkg, documentPart, 'VisioDocument');
	const mastersPart = await related(pkg, documentPart, 'masters', false);
	if (!mastersPart) fail('EDIT_TARGET_NOT_FOUND', 'The stencil has no masters.');
	const entry = children(await visioXml(pkg, mastersPart, 'Masters'), 'Master').find(
		(node) => attribute(node, 'ID') === stencilMasterId,
	);
	if (!entry) fail('EDIT_TARGET_NOT_FOUND', 'The stencil has no such master.');
	const part = await indexedPart(pkg, mastersPart, entry, 'master');
	const media = await masterPictures(pkg, part);
	const contents = await visioXml(pkg, part, 'MasterContents');
	const name = attribute(entry, 'NameU') ?? attribute(entry, 'Name') ?? `Master ${stencilMasterId}`;
	return addMasterToDrawing(
		source,
		{
			uniqueId: attribute(entry, 'UniqueID') ?? '',
			name,
			entry(masters, id, unique, relationshipId) {
				const node = masters.ownerDocument!.importNode(entry, true) as Element;
				if (node.namespaceURI !== masters.namespaceURI)
					fail('UNSUPPORTED_EDIT_PACKAGE', 'The stencil uses another Visio namespace.');
				node.setAttribute('ID', id);
				node.setAttribute('NameU', unique);
				if (node.hasAttribute('Name') && attribute(entry, 'Name') === name)
					node.setAttribute('Name', unique);
				// The master page's style references name the stencil's styles; the drawing's own
				// default applies instead.
				for (const sheet of children(node, 'PageSheet'))
					for (const style of STYLE_ATTRIBUTES)
						if (sheet.hasAttribute(style)) sheet.setAttribute(style, '0');
				for (const rel of children(node, 'Rel')) rel.setAttributeNS(NS.r, 'r:id', relationshipId);
				masters.appendChild(node);
			},
			media,
			contents(document, touch) {
				const map = styleMapper(stencilDocument, document, touch);
				const root = copy(contents);
				if (root.namespaceURI !== document.namespaceURI)
					fail('UNSUPPORTED_EDIT_PACKAGE', 'The stencil uses another Visio namespace.');
				for (const shape of Array.from(root.getElementsByTagNameNS(root.namespaceURI, 'Shape')))
					map(shape);
				return root;
			},
		},
		limits,
		maxOutput,
		deadline,
		check,
	);
}
