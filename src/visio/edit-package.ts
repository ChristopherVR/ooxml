import JSZip from 'jszip';
import { elements } from '../xml/index.js';
import { VisioPackage } from './package.js';
import { fail, decodePath, type VisioPackageLimits } from './package-common.js';
import { related, indexedPart, visioXml } from './parts.js';
import { attribute, children } from './sheet.js';

export async function openEditablePackage(
	bytes: Uint8Array,
	limits: VisioPackageLimits,
	check: () => void,
): Promise<{ pkg: VisioPackage; parts: Map<string, Uint8Array>; pages: Map<string, string> }> {
	const pkg = await VisioPackage.open(bytes, limits);
	const types = await pkg.readXml('[Content_Types].xml', 'Types');
	if (types.namespaceURI !== 'http://schemas.openxmlformats.org/package/2006/content-types')
		fail('INVALID_CONTENT_TYPES', 'Invalid content type namespace.');
	for (const node of elements(types))
		if (
			node.namespaceURI !== types.namespaceURI ||
			!['Default', 'Override'].includes(node.localName)
		)
			fail('INVALID_CONTENT_TYPES', 'Invalid content type element.');
	const documentPart = await related(pkg, '', 'document');
	if (!documentPart) fail('INVALID_DOCUMENT', 'Missing Visio document part.');
	const overrides = elements(types).filter(
		(node) => node.localName === 'Override' && attribute(node, 'PartName') === `/${documentPart}`,
	);
	if (
		overrides.length !== 1 ||
		overrides[0]?.getAttribute('ContentType') !== 'application/vnd.ms-visio.drawing.main+xml'
	)
		fail('UNSUPPORTED_FORMAT', 'Editing supports only VSDX drawing packages.');
	for (const node of elements(types)) {
		const type = node.getAttribute('ContentType') ?? '';
		if (/digital-signature|macroEnabled|vbaProject/i.test(type))
			fail('UNSUPPORTED_EDIT_PACKAGE', 'Signed or macro-enabled packages cannot be edited.');
	}
	const parts = new Map<string, Uint8Array>();
	for (const path of pkg.paths()) {
		check();
		if (/(^|\/)_xmlsignatures(\/|$)|(^|\/)vbaProject\.bin$/i.test(decodePath(path)))
			fail('UNSUPPORTED_EDIT_PACKAGE', 'Signed packages cannot be edited.');
		if (path === '_rels/.rels' || /(^|\/)\_rels\/[^/]+\.rels$/.test(path)) {
			const source =
				path === '_rels/.rels' ? '' : path.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, '$1$2');
			for (const rel of (await pkg.relationships(source)).values())
				if (/digital-signature|vbaProject/i.test(rel.type))
					fail('UNSUPPORTED_EDIT_PACKAGE', 'Signed or macro-enabled packages cannot be edited.');
		}
		// Validate size and CRC even for unused opaque parts; never fetch external targets.
		parts.set(path, new Uint8Array(await pkg.readBytes(path)));
	}
	await visioXml(pkg, documentPart, 'VisioDocument');
	const pagesPart = await related(pkg, documentPart, 'pages');
	if (!pagesPart) fail('INVALID_DOCUMENT', 'Missing pages part.');
	const pages = new Map<string, string>();
	const paths = new Set<string>();
	for (const page of children(await visioXml(pkg, pagesPart, 'Pages'), 'Page')) {
		check();
		const id = attribute(page, 'ID');
		if (!id || pages.has(id)) fail('INVALID_PAGE_ID', 'Page IDs must be present and unique.');
		const rels = children(page, 'Rel');
		if (
			rels.length !== 1 ||
			!rels[0]!.getAttributeNS(
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'id',
			)
		)
			fail('INVALID_RELATIONSHIP', 'Page requires one namespace-qualified relationship ID.');
		const path = await indexedPart(pkg, pagesPart, page, 'page');
		if (paths.has(path)) fail('INVALID_PAGE_ID', 'Multiple pages reference the same page part.');
		paths.add(path);
		pages.set(id, path);
	}
	return { pkg, parts, pages };
}

/** Bounded emitted output. ZIP metadata/compressed representation is not preserved. */
export function writeEditedPackage(
	parts: ReadonlyMap<string, Uint8Array>,
	maxBytes: number,
	deadline: number,
	check: () => void,
): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [path, bytes] of parts) {
		check();
		zip.file(path, bytes, { createFolders: false });
	}
	return new Promise((resolve, reject) => {
		const stream = zip.generateInternalStream({
			type: 'uint8array',
			compression: 'STORE',
			streamFiles: true,
		});
		let chunks: Uint8Array[] = [];
		let size = 0;
		let settled = false;
		const abort = (error: unknown) => {
			if (settled) return;
			settled = true;
			stream.pause();
			chunks = [];
			clearTimeout(timer);
			reject(error);
		};
		const timer = setTimeout(
			() => {
				try {
					fail('LIMIT_RUNTIME', 'Visio edit deadline exceeded.');
				} catch (error) {
					abort(error);
				}
			},
			Math.max(1, deadline - Date.now()),
		);
		stream
			.on('data', (chunk: Uint8Array) => {
				if (settled) return;
				try {
					check();
					size += chunk.length;
					if (size > maxBytes) fail('LIMIT_EDIT_OUTPUT', 'Saved package exceeds output limit.');
					chunks.push(chunk);
				} catch (error) {
					abort(error);
				}
			})
			.on('error', abort)
			.on('end', () => {
				if (settled) return;
				try {
					check();
					const bytes = new Uint8Array(size);
					let offset = 0;
					for (const chunk of chunks) {
						bytes.set(chunk, offset);
						offset += chunk.length;
					}
					settled = true;
					clearTimeout(timer);
					chunks = [];
					resolve(bytes);
				} catch (error) {
					abort(error);
				}
			})
			.resume();
	});
}
